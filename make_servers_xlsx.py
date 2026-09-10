#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Regenerate servers.xlsx - the Excel server list read by bssv_monitor.py.

    python tools\\make_servers_xlsx.py            # writes ..\\servers.xlsx
    python tools\\make_servers_xlsx.py D:\\x.xlsx  # writes to a chosen path

Existing files are NOT overwritten unless --force is passed, so you cannot
lose your server list by accident.
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ENVIRONMENTS = ["PROD", "PRE-PROD", "UAT", "TEST", "DEV", "DR", "TRAINING"]

HEADERS = [
    ("Enabled",        10, "Y = this row is checked, N = skipped."),
    ("ServerName",     22, "Friendly name shown on the dashboard and in alerts."),
    ("Environment",    14, "REQUIRED. Pick from the dropdown (PROD / PRE-PROD / UAT / TEST / "
                           "DEV / DR / TRAINING) or type your own. Drives the dashboard "
                           "environment selector, the alert subject line and per-environment "
                           "e-mail routing in config.ini [email.recipients]."),
    ("ServiceName",    28, "BSSV service being probed, e.g. RI_AddressBookManager."),
    ("EndpointURL",    62, "Full https URL the SOAP payload is POSTed to."),
    ("PayloadPath",    46, "XML file. Relative paths resolve against payload_dir in config.ini."),
    ("SOAPAction",     16, 'SOAPAction header. BSSV normally uses "" (two quotes).'),
    ("ExpectedText",   30, "Optional. Response MUST contain this string to count as success."),
    ("TimeoutSeconds", 16, "Optional. Overrides [request] timeout_seconds."),
    ("VerifySSL",      12, "Optional Y/N. Overrides [request] verify_ssl."),
    ("AlertEmailTo",   34, "Optional. Overrides [email] to_addresses for this row."),
    ("Notes",          34, "Free text - ignored by the program."),
]

SAMPLE_ROWS = [
    ["Y", "BSSV-PROD", "PROD", "RI_AddressBookManager",
     "https://bssv-prod-host.example.com:7851/PD920OR/RI_AddressBookManager",
     "RI_AddressBookManager_getAddressBook.xml", '""', "getAddressBookResponse",
     45, "N", "jde-support@yourcompany.com", "Production - alerts to the support DL"],
    ["Y", "BSSV-UAT", "UAT", "RI_AddressBookManager",
     "https://bssv-uat-host.example.com:7451/PY920/RI_AddressBookManager",
     "RI_AddressBookManager_getAddressBook.xml", '""', "getAddressBookResponse",
     45, "N", "", "UAT - alerts fall back to [email.recipients] UAT"],
    ["Y", "BSSV-DEV", "DEV", "RI_AddressBookManager",
     "https://bssv-dev-host.example.com:7253/DV920/RI_AddressBookManager",
     "RI_AddressBookManager_getAddressBook.xml", '""', "getAddressBookResponse",
     30, "N", "", "Dev sanity check"],
    ["N", "BSSV-DR", "DR", "RI_AddressBookManager",
     "https://bssv-dr-host.example.com:7851/PD920OR/RI_AddressBookManager",
     "RI_AddressBookManager_getAddressBook.xml", '""', "", 45, "N", "",
     "Disabled until DR go-live"],
]

HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build(path: str) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Servers"

    for col, (name, width, note) in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        cell.comment = None
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 28

    for r, row in enumerate(SAMPLE_ROWS, start=2):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (5, 6, 12)))

    yes_no = DataValidation(type="list", formula1='"Y,N"', allow_blank=True,
                            showDropDown=False, promptTitle="Y or N",
                            prompt="Y = yes, N = no")
    ws.add_data_validation(yes_no)
    yes_no.add("A2:A2000")
    yes_no.add("J2:J2000")

    # Environment dropdown. allow_blank/showErrorMessage stay off so a site can
    # type an environment code of its own without Excel rejecting the cell.
    env = DataValidation(
        type="list", formula1='"' + ",".join(ENVIRONMENTS) + '"',
        allow_blank=True, showDropDown=False, showErrorMessage=False,
        promptTitle="Environment",
        prompt="Pick one, or type your own code. Required - it drives the "
               "dashboard filter, the alert subject and e-mail routing.",
    )
    ws.add_data_validation(env)
    env.add("C2:C2000")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1000"

    # ---- documentation sheet ----------------------------------------------
    doc = wb.create_sheet("ReadMe")
    doc.column_dimensions["A"].width = 20
    doc.column_dimensions["B"].width = 100
    doc["A1"] = "Column"
    doc["B1"] = "Meaning"
    for cell in (doc["A1"], doc["B1"]):
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
    for r, (name, _w, note) in enumerate(HEADERS, start=2):
        doc.cell(row=r, column=1, value=name).font = Font(bold=True)
        doc.cell(row=r, column=2, value=note).alignment = Alignment(wrap_text=True, vertical="top")

    r = len(HEADERS) + 3
    doc.cell(row=r, column=1, value="Success rule").font = Font(bold=True, color="C00000")
    doc.cell(
        row=r, column=2,
        value=("A check counts as SUCCESS only when: HTTP 200  AND  the response contains no "
               "SOAP Fault  AND  (if ExpectedText is filled in) the response contains that text. "
               "Anything else - non-200, timeout, TLS error, connection refused, SOAP Fault, "
               "missing ExpectedText - is FAILED and triggers an e-mail alert."),
    ).alignment = Alignment(wrap_text=True, vertical="top")

    r += 2
    doc.cell(row=r, column=1, value="Environment").font = Font(bold=True, color="1F3864")
    doc.cell(
        row=r, column=2,
        value=("Environment is a REQUIRED column - a row with a blank Environment is skipped. "
               "It flows through everything: the CSV log, the dashboard environment selector "
               "(counters, chart, server grid and history all follow the selection), the alert "
               "subject line, and the grouping inside the alert e-mail. config.ini can also "
               "route alerts per environment via the [email.recipients] section, e.g. "
               "PROD = ops-oncall@company.com and DEV = jde-team@company.com. Recipient "
               "precedence: the AlertEmailTo cell, then [email.recipients], then "
               "[email] to_addresses."),
    ).alignment = Alignment(wrap_text=True, vertical="top")

    r += 2
    doc.cell(row=r, column=1, value="Credentials").font = Font(bold=True)
    doc.cell(
        row=r, column=2,
        value=("Credentials live inside each XML payload (wsse:UsernameToken). Keep one XML per "
               "server/environment and restrict NTFS permissions on the payloads folder."),
    ).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(path)
    print(f"Written: {path}")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--force"]
    force = "--force" in sys.argv
    default = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "servers.xlsx")
    target = args[0] if args else default
    if os.path.exists(target) and not force:
        sys.exit(f"Refusing to overwrite existing file: {target}\nPass --force if you really mean it.")
    build(target)
