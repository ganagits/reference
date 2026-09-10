# BSSV Monitor

Scheduled connectivity + health monitoring for Oracle JD Edwards **Business
Services (BSSV)** endpoints.

A Python program POSTs a SOAP payload to each configured BSSV server, decides
whether the response is healthy, writes every result to a rolling CSV,
republishes a 72-hour JSON feed, and — **only when a check fails** — sends an
e-mail alert through a PowerShell `Send-MailMessage` script (no stored
password). A static, self-refreshing dashboard hosted in IIS or XAMPP shows the
last 72 hours.

```
   servers.xlsx ─┐
   payloads\*.xml ├──► bssv_monitor.py ──► logs\bssv_monitor_log.csv  (72h, auto-purged)
   config.ini ────┘          │                     │
                             │                     └──► <web_root>\status.json ──► dashboard
                             └── on failure ──► send_alert.ps1 ──► Send-MailMessage
```

---

## 1. What counts as success

A check is **SUCCESS** only when all three hold:

1. HTTP status is **200**, **and**
2. the response body contains **no SOAP Fault**, **and**
3. if `ExpectedText` is filled in for that row, the response **contains that
   text**.

Everything else is **FAILED** and is recorded with a machine-readable reason:

| `failure_reason` | Meaning |
|---|---|
| `HTTP_STATUS` | Server answered with a non-200 code (401, 404, 500 …) |
| `SOAP_FAULT` | HTTP 200 but the body carries a `soap:Fault` — bad credentials, E1 error, service not deployed |
| `EXPECTED_TEXT_MISSING` | Answered 200 with no fault, but the `ExpectedText` marker was absent |
| `CONNECT_TIMEOUT` | TCP connect did not complete in time |
| `READ_TIMEOUT` | Connected, but no response within the timeout |
| `CONNECTION_ERROR` | Refused / host unreachable / DNS failure |
| `SSL_ERROR` | Certificate or TLS handshake problem |
| `PAYLOAD_UNREADABLE` | The XML file in `PayloadPath` is missing or locked |
| `UNEXPECTED_ERROR` | Anything else — full detail in `logs\bssv_monitor.log` |

Successes are logged and charted; they are **never** e-mailed.

---

## 2. Folder layout

```
bssv-monitor\
├─ bssv_monitor.py           the monitor
├─ config.ini                every tunable setting
├─ servers.xlsx              the server list (Excel)
├─ send_alert.ps1            PowerShell mailer (Send-MailMessage, no password)
├─ requirements.txt          Python dependencies
├─ run_once.bat              single pass  - for Windows Task Scheduler
├─ run_loop.bat              resident loop - for a console window or NSSM
├─ payloads\
│   └─ RI_AddressBookManager_getAddressBook.xml
├─ dashboard\
│   └─ index.html            published to web_root on every pass
├─ tools\
│   └─ make_servers_xlsx.py  regenerates a blank servers.xlsx
└─ logs\
    ├─ bssv_monitor_log.csv  rolling 72h result log (auto-purged)
    ├─ bssv_monitor.log      application log (rotating)
    ├─ alert_state.json      alert throttle state
    ├─ archive\              purged rows, one CSV per month
    └─ responses\            raw bodies of failed responses
```

---

## 3. Install

Windows Server / Windows 10+ with Python 3.8 or newer.

```bat
cd C:\BSSV\bssv-monitor
python -m pip install -r requirements.txt
```

Dependencies: `requests`, `openpyxl`, `urllib3`. Nothing else — no database, no
web framework, no scheduler service.

If the machine has no internet access, download the three wheels on a connected
machine and `pip install <wheel>` them locally.

---

## 4. The payload XML

The payloads folder holds one SOAP envelope per server/service. The supplied
sample was extracted from `SAMPLE\PROJ-soapui-project.xml` — the `getAddressBook`
request of `RI_AddressBookManager`:

```xml
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                  xmlns:orac="http://oracle.e1.bssv.JPR01000/">
  <soapenv:Header>
    <wsse:Security xmlns:wsse="...wssecurity-secext-1.0.xsd" soapenv:mustUnderstand="1">
      <wsse:UsernameToken>
        <wsse:Username>DUMMY_USER</wsse:Username>
        <wsse:Password Type="...#PasswordText">DUMMY_PASSWORD</wsse:Password>
      </wsse:UsernameToken>
    </wsse:Security>
  </soapenv:Header>
  <soapenv:Body>
    <orac:getAddressBook>
      <entity><entityId>1</entityId></entity>
    </orac:getAddressBook>
  </soapenv:Body>
</soapenv:Envelope>
```

**Credentials are hardcoded inside each XML file.** That is deliberate — it
keeps one file self-contained per environment. Practical consequences:

* Use one payload file per server/environment when the credentials differ
  (`..._DEV.xml`, `..._PROD.xml`) and point each Excel row at its own file.
* Use a **dedicated, low-privilege monitoring E1 user**, not a named person's
  account — a monitor that runs every 5 minutes will lock out an account with an
  expiring password.
* Restrict NTFS permissions on `payloads\` to the service account that runs the
  monitor.
* `getAddressBook` for `entityId 1` is a good probe: read-only, always exists,
  cheap. Avoid `addAddressBook` and any other write operation.

To add a different service, export the request from SoapUI (right-click the
request → *Copy*), save it as an `.xml` file in `payloads\`, and add a row to
the Excel.

---

## 5. `servers.xlsx` — the server list

Sheet **`Servers`**, row 1 is the header. Column order does not matter; header
names do. A second sheet, `ReadMe`, repeats this table inside the workbook.

| Column | Required | Meaning |
|---|---|---|
| `Enabled` | yes | `Y` = check this row, `N` = skip it |
| `ServerName` | yes | Friendly name shown on the dashboard and in alerts |
| `Environment` | **yes** | `PROD` / `PRE-PROD` / `UAT` / `TEST` / `DEV` / `DR` / `TRAINING` — a dropdown, but you can type your own code. See §5.1 |
| `ServiceName` | no | BSSV service being probed, e.g. `RI_AddressBookManager` |
| `EndpointURL` | yes | Full URL the payload is POSTed to |
| `PayloadPath` | yes | XML file; relative paths resolve against `payload_dir` |
| `SOAPAction` | no | SOAPAction header. BSSV normally uses `""` (two quote characters) |
| `ExpectedText` | no | If filled in, the response **must** contain this string |
| `TimeoutSeconds` | no | Overrides `[request] timeout_seconds` for this row |
| `VerifySSL` | no | `Y`/`N`, overrides `[request] verify_ssl` for this row |
| `AlertEmailTo` | no | Overrides `[email] to_addresses` for this row |
| `Notes` | no | Free text — ignored by the program |

Sample rows shipped in the workbook:

| Enabled | ServerName | Environment | ServiceName | EndpointURL | ExpectedText |
|---|---|---|---|---|---|
| Y | BSSV-PROD | PROD | RI_AddressBookManager | `https://bssv-prod-host:7851/PD920OR/RI_AddressBookManager` | `getAddressBookResponse` |
| Y | BSSV-UAT | UAT | RI_AddressBookManager | `https://bssv-uat-host:7451/PY920/RI_AddressBookManager` | `getAddressBookResponse` |
| Y | BSSV-DEV | DEV | RI_AddressBookManager | `https://bssv-dev-host:7253/DV920/RI_AddressBookManager` | `getAddressBookResponse` |
| N | BSSV-DR | DR | RI_AddressBookManager | `https://bssv-dr-host:7851/PD920OR/RI_AddressBookManager` | |

`getAddressBookResponse` is the recommended `ExpectedText` for this probe: it is
the response element name, so its presence proves E1 actually executed the
business function rather than merely accepting the connection.

Edit the workbook in Excel and save — the next pass picks up the change, no
restart needed. To start from a clean workbook:

```bat
python tools\make_servers_xlsx.py --force
```

### 5.1 Environment is a first-class dimension

Not every server is production, and the tool never treats them as if they were.
`Environment` is **required** — a row with a blank Environment is skipped with a
warning — and the value is upper-cased, so `prod`, `Prod` and `PROD` are the same
environment. It then flows through every part of the system:

| Where | What the environment does |
|---|---|
| `servers.xlsx` | Dropdown of the seven standard codes; type your own if your site uses different ones |
| CSV log | `environment` column on every row |
| **Dashboard** | An environment selector across the top. Picking one rescopes the whole page — counters, chart, server grid and history. Server cards are grouped under an environment heading, and every card and history row carries an environment tag; `PROD` is tinted red |
| **Alert subject** | Leads with the environment: `[BSSV ALERT] PROD - 2 BSSV check(s) FAILED - BSSV-PROD-1, BSSV-PROD-2` — so an Outlook rule can act on PROD alone |
| **Alert body** | Failures grouped under an environment banner, production flagged explicitly |
| **Alert routing** | `[email.recipients]` in `config.ini` sends each environment's alerts to a different distribution list |
| Ordering | Production sorts first everywhere — server cards, environment chips, and the order alert mails are sent |

Environments recognised as production for the red tint and the "PRODUCTION"
flag: `PROD`, `PRODUCTION`, `PRD`, `LIVE`. The sort order used throughout is
PROD → PRE-PROD → DR → UAT → TEST → QA → DEV → TRAINING, with anything unknown
after that, alphabetically.

---

## 6. `config.ini`

Grouped by section. Paths may be absolute or relative to the folder holding
`config.ini`.

### `[general]`

| Key | Default | Meaning |
|---|---|---|
| `run_interval_seconds` | `300` | **Frequency.** Seconds between passes in loop mode |
| `run_mode` | `loop` | `loop` = stay resident; `once` = single pass and exit |
| `skip_missed_ticks` | `true` | If a pass overruns the interval, skip the missed tick instead of running back-to-back |
| `excel_config` | `servers.xlsx` | Workbook holding the server list |
| `excel_sheet` | `Servers` | Sheet inside it |
| `payload_dir` | `payloads` | Base folder for relative `PayloadPath` values |
| `max_workers` | `4` | Servers probed in parallel; `1` = strictly sequential |

### `[request]`

| Key | Default | Meaning |
|---|---|---|
| `timeout_seconds` | `30` | Per-request timeout |
| `retries` | `1` | Extra attempts after the first failure. A check is FAILED only after all attempts fail — this suppresses alerts from single dropped packets |
| `retry_delay_seconds` | `5` | Pause between attempts |
| `verify_ssl` | `false` | TLS certificate verification |
| `ca_bundle` | *(blank)* | PEM file for your internal CA; used when `verify_ssl = true` |
| `default_soap_action` | `""` | SOAPAction header when the Excel cell is blank |
| `content_type` | `text/xml; charset=utf-8` | SOAP 1.1 content type |
| `user_agent` | `BSSV-Monitor/1.0` | Handy for spotting the monitor in WebLogic access logs |
| `http_proxy` / `https_proxy` | *(blank)* | Leave blank for a direct connection |

**On `verify_ssl`.** BSSV usually presents a certificate signed by an internal
CA, which Python does not trust out of the box. Shipping default is `false` so
the monitor works immediately. To do it properly: export the internal root CA as
Base-64 `.cer`, rename to `.pem`, then set `verify_ssl = true` and
`ca_bundle = C:\certs\internal-ca.pem`. With `verify_ssl = false` an expired or
swapped certificate will **not** be detected.

### `[logging]`

| Key | Default | Meaning |
|---|---|---|
| `result_log` | `logs\bssv_monitor_log.csv` | Rolling result log (opens directly in Excel) |
| `retention_hours` | `72` | **Autopurge window.** Rows older than this are removed after every pass |
| `archive_purged` | `true` | Purged rows are appended to a monthly archive instead of being lost |
| `archive_dir` | `logs\archive` | `bssv_monitor_log_YYYYMM.csv` per month |
| `app_log` | `logs\bssv_monitor.log` | Rotating application log |
| `app_log_level` | `INFO` | `DEBUG` / `INFO` / `WARNING` / `ERROR` |
| `app_log_max_bytes` / `app_log_backup_count` | `5 MB` / `5` | Rotation |
| `save_failed_responses` | `true` | Keep the raw body of failed responses for troubleshooting |
| `failed_response_dir` | `logs\responses` | Where those bodies land |
| `failed_response_retention_hours` | `72` | Those files are auto-purged too |

Result-log columns:

`run_id, timestamp_local, timestamp_utc, server_name, environment,
service_name, endpoint_url, payload_file, status, http_status, response_ms,
failure_reason, detail, attempts, email_sent`

Timestamps are written twice: `timestamp_local` for reading in Excel,
`timestamp_utc` for the arithmetic (purge window, chart buckets) — so the log
stays correct across a DST change.

### `[dashboard]`

| Key | Default | Meaning |
|---|---|---|
| `enabled` | `true` | Publish the JSON feed |
| `web_root` | `C:\inetpub\wwwroot\bssv-monitor` | **Where the dashboard is hosted.** Created if missing |
| `json_file_name` | `status.json` | Feed name inside `web_root` |
| `window_hours` | `72` | Rolling window shown |
| `copy_dashboard_files` | `true` | Copy `dashboard\index.html` into `web_root` on every pass |
| `max_detail_rows` | `1500` | Cap on individual checks embedded in the feed (~350 KB per 1000 rows). Counters and chart always cover the full window |

The feed is written to a `.tmp` file and then atomically renamed, so a browser
refreshing mid-write never reads a half-written file.

### `[email]`

| Key | Default | Meaning |
|---|---|---|
| `enabled` | `true` | Master switch for alerting |
| `powershell_script` | `send_alert.ps1` | The mailer |
| `powershell_exe` | `powershell.exe` | Use `pwsh.exe` for PowerShell 7 |
| `smtp_server` / `smtp_port` | — / `25` | Your relay |
| `use_ssl` | `false` | STARTTLS (port 587 typically) |
| `from_address` | — | Envelope sender; must be allowed to relay |
| `to_addresses` | — | Comma separated **default** recipients (fallback) |
| `cc_addresses` | *(blank)* | Comma separated |
| `subject_prefix` | `[BSSV ALERT]` | Prefix for mail rules; the environment follows it |
| `throttle_minutes` | `30` | No second alert for the same server within this window. `0` = alert every pass |
| `consolidate` | `true` | One mail listing all failures in a pass, instead of one mail each |
| `group_by_environment` | `true` | With `consolidate`, send a **separate mail per environment** so PROD keeps its own thread |
| `attach_response` | `true` | Attach the failed response body |
| `send_recovery` | `false` | Send one "RECOVERED" mail when a previously-alerting server comes back |

### `[email.recipients]` — per-environment routing

A DEV outage at 2am should not page the people who own production. Map each
`Environment` value to its own recipients:

```ini
[email.recipients]
PROD      = ops-oncall@yourcompany.com, jde-support@yourcompany.com
PRE-PROD  = jde-support@yourcompany.com
UAT       = jde-test-team@yourcompany.com
TEST      = jde-test-team@yourcompany.com
DEV       = jde-dev-team@yourcompany.com
DR        = infra-team@yourcompany.com
```

Keys are case-insensitive and every environment is optional. **Recipient
precedence, highest first:**

1. the `AlertEmailTo` cell on the Excel row,
2. the matching `[email.recipients]` entry,
3. `[email] to_addresses`.

With `group_by_environment = true` (the default) a pass that breaks PROD, UAT
and DEV at once sends three separate mails — each to its own list, each with its
environment leading the subject — instead of one mail that everybody has to
read.

**Throttling matters.** With a 5-minute interval and no throttle, one server
down overnight produces ~96 e-mails. The default sends one, then one every 30
minutes while it stays down. The throttle is tracked per server, so a PROD
outage never masks a UAT one.

No password is stored anywhere. `send_alert.ps1` calls `Send-MailMessage`
without `-Credential`, so the SMTP relay must accept anonymous submission from
the monitoring host — ask your mail team to allow that host's IP on the internal
relay.

---

## 7. Running it

**Task Scheduler (recommended for production).** Set `run_mode = once` in
`config.ini`, then:

```
Program/script : C:\BSSV\bssv-monitor\run_once.bat
Start in       : C:\BSSV\bssv-monitor
Trigger        : Daily, repeat every 5 minutes, for a duration of 1 day
Settings       : Run whether user is logged on or not
                 If the task is already running, do not start a new instance
```

Task Scheduler survives reboots and restarts a crashed pass. This is the safer
choice for an unattended server.

**Resident loop.** Set `run_mode = loop` and run `run_loop.bat`. Useful during
setup because everything is visible on the console. To make it a real service,
wrap it with [NSSM](https://nssm.cc):

```bat
nssm install BSSVMonitor "C:\BSSV\bssv-monitor\run_loop.bat"
nssm set BSSVMonitor AppDirectory C:\BSSV\bssv-monitor
nssm start BSSVMonitor
```

**Command-line flags** (they override `config.ini`):

```bat
python bssv_monitor.py --once            :: one pass, then exit
python bssv_monitor.py --loop            :: stay resident
python bssv_monitor.py --once --dry-run  :: probe and log, never send e-mail
python bssv_monitor.py --once --verbose  :: debug output on the console
python bssv_monitor.py --test-email      :: send a test alert and exit
python bssv_monitor.py --config D:\alt\config.ini
```

---

## 8. Hosting the dashboard

The dashboard is a **single static HTML file** that reads `status.json` from the
same folder. No PHP, no ASP.NET, no backend — it works identically in IIS and
XAMPP. It refreshes itself every **60 seconds** and shows the rolling
**72 hours**:

* an **environment selector** across the top — `All` plus one chip per
  environment, each showing its check count and a red dot when a server in it is
  down. Picking one rescopes the entire page: counters, chart, server grid and
  history. The choice is remembered in the browser, so a screen mounted on the
  wall can sit permanently on PROD;
* All / Success / Failed counters and the success rate for the selected scope;
* an hourly success-vs-failure chart, redrawn per environment;
* a server grid **grouped by environment**, production first, with uptime %,
  check count, failure count and average response time per server;
* the full check history **newest first**, with All / Success / Failed filters,
  a search box, and an environment tag on every row.

### IIS

1. Set `web_root = C:\inetpub\wwwroot\bssv-monitor` in `config.ini`.
2. Run one pass — the folder, `index.html` and `status.json` are created.
3. IIS Manager → *Sites* → *Default Web Site* → **Add Application**
   (alias `bssv-monitor`, physical path as above), or just browse the folder
   under the default site.
4. Make sure the `.json` MIME type exists: IIS Manager → the site → **MIME
   Types** → add `.json` = `application/json` if it is missing.
5. Browse to `http://<server>/bssv-monitor/`.

Give the account running the monitor **write** permission on that folder, and
`IIS_IUSRS` **read**.

### XAMPP

1. Set `web_root = C:\xampp\htdocs\bssv-monitor`.
2. Run one pass, start Apache, browse to
   `http://localhost/bssv-monitor/`.

### Anywhere else

Any static web server will do. The one thing that does **not** work is opening
`index.html` from disk with a `file://` URL — browsers block the `fetch` of
`status.json` from the filesystem. Serve it over HTTP.

---

## 9. Verifying the setup

Work through these in order; each one isolates a different link in the chain.

```bat
:: 1. Is the Excel readable and are the endpoints reachable? No mail sent.
python bssv_monitor.py --once --dry-run --verbose

:: 2. Does the SMTP relay accept our mail?
python bssv_monitor.py --test-email

:: 3. Full pass with alerting live.
python bssv_monitor.py --once
```

Then check:

* `logs\bssv_monitor_log.csv` has one row per enabled server.
* `<web_root>\status.json` exists and its `generated_at_local` is recent.
* The dashboard loads over HTTP and the counters match the CSV.

A quick negative test: add a row pointing at a deliberately wrong port, run one
pass, and confirm the alert arrives and the row shows up red on the dashboard.

---

## 10. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Excel config not found` | `excel_config` path is wrong, or the workbook is open with an exclusive lock. Close it in Excel |
| `missing required column(s)` | A header was renamed. `ServerName`, `Environment`, `EndpointURL`, `PayloadPath` are mandatory |
| `Excel row N skipped - ServerName/Environment/EndpointURL/PayloadPath incomplete` | One of those four cells is blank on that row — most often a missing `Environment` |
| A server is missing from the dashboard | Its environment is not the one currently selected. Click **All** in the environment bar |
| The wrong team got the alert | Check the precedence in §6: an `AlertEmailTo` cell beats `[email.recipients]`, which beats `to_addresses`. Also check the environment spelling — `PRD` and `PROD` are different environments |
| Every server reports `SOAP_FAULT` with "Invalid user or password" | The credentials in the payload XML are wrong or the E1 account is locked/expired |
| `SSL_ERROR` on every server | Internal CA not trusted. Either set `ca_bundle`, or set `verify_ssl = false` |
| `CONNECTION_ERROR` from the monitoring host only | Firewall between the host and the BSSV port (7251/7253/7851 …). Test with `Test-NetConnection <host> -Port 7851` |
| `HTTP_STATUS 404` | Wrong path segment — the environment code (`DV920`, `PD920OR`) or the service name is wrong in `EndpointURL` |
| `EXPECTED_TEXT_MISSING` but the service looks fine | The response element name differs for this operation. Blank the cell, run once, and read the saved body in `logs\responses\` to pick the right marker |
| No alert e-mail, log says `MAIL FAILED` | Relay rejected the message. Test by hand: `Send-MailMessage -SmtpServer <relay> -From <from> -To <you> -Subject test -Body test` |
| No alert and no error | Throttled. Check `logs\alert_state.json`, or set `throttle_minutes = 0` while testing |
| `PowerShell executable not found` | Set `powershell_exe` to the full path, e.g. `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe` |
| Dashboard shows "Could not load status.json" | The feed is not in `web_root`, the site is not serving `.json`, or you opened the file with `file://` |
| Dashboard never updates | The monitor is not running, or `web_root` points somewhere the web server does not serve |
| CSV keeps growing past 72 hours | `retention_hours` was raised, or the file is held open by Excel so the rewrite fails — check `logs\bssv_monitor.log` |

---

## 11. Operating notes

* **Interval.** 5 minutes (`300`) is a sensible default: fast enough to catch an
  outage, gentle enough on a production E1 server. Below 60 seconds you are
  load-testing your own BSSV instance.
* **Retention.** 72 hours is the dashboard window. Raise `retention_hours` and
  `window_hours` together if you want a longer view, and keep an eye on
  `max_detail_rows` — the feed is re-downloaded by every open browser tab every
  60 seconds.
* **Archive.** With `archive_purged = true` nothing is actually lost: the
  monthly CSVs under `logs\archive\` hold the full history for trend reporting.
* **Clock.** Chart buckets and the purge window use UTC internally and display
  local time, so a DST shift does not distort the 72-hour window.
* **Concurrency.** `max_workers = 4` probes four servers at once. Set it to `1`
  if your BSSV instance is sensitive to parallel sessions.
* **Read-only probes only.** Never point a payload at a write operation
  (`addAddressBook`, order entry …) — the monitor runs unattended, forever.

---

## 12. Extending it

* **New service:** export the request from SoapUI into `payloads\`, add an Excel
  row, done. No code change.
* **New environment:** copy the payload, change the credentials inside it, add a
  row with the new `Environment` and `EndpointURL`. The environment appears in
  the dashboard selector and the alert subject on the next pass; add an
  `[email.recipients]` line if it needs its own distribution list. The seven
  codes in the dropdown are only a convenience — typing `PY920` or `JPD` works
  exactly the same, it just will not get the production tint or the top sort
  slot unless it is one of `PROD` / `PRODUCTION` / `PRD` / `LIVE`.
* **Different success rule:** the entire rule lives in `evaluate()` in
  `bssv_monitor.py` (about 15 lines).
* **Different alert channel:** `send_alert.ps1` is called with plain arguments
  (`-SmtpServer -Port -From -To -Subject -BodyFile`). Swap its body for a Teams
  webhook or a ticket-system call and the Python side does not change.
