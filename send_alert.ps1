<#
.SYNOPSIS
    Sends a BSSV Monitor alert e-mail using Send-MailMessage (no credentials).

.DESCRIPTION
    Called by bssv_monitor.py whenever one or more BSSV checks fail.
    Uses an anonymous/relay SMTP connection - no password is stored anywhere.
    The HTML body is passed as a file so that quoting and non-ASCII characters
    survive the command line intact.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File send_alert.ps1 `
        -SmtpServer smtp.yourcompany.com -Port 25 `
        -From bssv-monitor@yourcompany.com -To "a@x.com,b@x.com" `
        -Subject "[BSSV ALERT] PROD failed" -BodyFile C:\temp\body.html

.NOTES
    Exit codes:  0 = sent,  1 = failed (message written to stderr).
    Send-MailMessage is deprecated but retained here because it is the
    zero-dependency option available on every supported Windows Server.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $SmtpServer,
    [int]    $Port = 25,
    [Parameter(Mandatory = $true)][string] $From,
    [Parameter(Mandatory = $true)][string] $To,          # comma or semicolon separated
    [string] $Cc = "",
    [Parameter(Mandatory = $true)][string] $Subject,
    [string] $BodyFile = "",                             # HTML body file (UTF-8)
    [string] $Body = "",                                 # inline body (used if BodyFile is empty)
    [string] $Attachments = "",                          # semicolon separated file paths
    [switch] $UseSsl,
    [ValidateSet("High", "Normal", "Low")]
    [string] $Priority = "High"
)

$ErrorActionPreference = "Stop"

function Split-AddressList {
    param([string] $Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    return @($Raw -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
}

try {
    # ---- recipients --------------------------------------------------------
    $toList = Split-AddressList $To
    if ($toList.Count -eq 0) { throw "No valid recipient in -To." }
    $ccList = Split-AddressList $Cc

    # ---- body --------------------------------------------------------------
    $htmlBody = $Body
    if (-not [string]::IsNullOrWhiteSpace($BodyFile)) {
        if (-not (Test-Path -LiteralPath $BodyFile)) { throw "BodyFile not found: $BodyFile" }
        $htmlBody = Get-Content -LiteralPath $BodyFile -Raw -Encoding UTF8
    }
    if ([string]::IsNullOrWhiteSpace($htmlBody)) { $htmlBody = "<p>(no body)</p>" }

    # ---- attachments (silently drop anything missing) ----------------------
    $attachList = @()
    foreach ($item in (Split-AddressList $Attachments)) {
        if (Test-Path -LiteralPath $item) { $attachList += $item }
    }

    # ---- assemble ----------------------------------------------------------
    $params = @{
        SmtpServer  = $SmtpServer
        Port        = $Port
        From        = $From
        To          = $toList
        Subject     = $Subject
        Body        = $htmlBody
        BodyAsHtml  = $true
        Priority    = $Priority
        Encoding    = [System.Text.Encoding]::UTF8
        ErrorAction = "Stop"
    }
    if ($ccList.Count -gt 0)    { $params["Cc"] = $ccList }
    if ($attachList.Count -gt 0){ $params["Attachments"] = $attachList }
    if ($UseSsl)                { $params["UseSsl"] = $true }

    # Send-MailMessage is obsolete in PS 6+; suppress the warning, keep behaviour.
    Send-MailMessage @params -WarningAction SilentlyContinue

    Write-Output "SENT to $($toList -join ', ') via $SmtpServer`:$Port"
    exit 0
}
catch {
    # Deliberately NOT Write-Error. Write-Error wraps the message in a block
    # showing this file, line and "CategoryInfo: NotSpecified", which buries
    # the one thing that matters - what the SMTP server actually said.
    $ex = $_.Exception

    $out = New-Object System.Collections.ArrayList
    [void]$out.Add("MAIL FAILED: " + $ex.Message)

    # SmtpException usually carries the real reason one level down.
    $inner = $ex.InnerException
    $depth = 0
    while ($inner -and $depth -lt 4) {
        [void]$out.Add("   caused by : " + $inner.Message)
        $inner = $inner.InnerException
        $depth++
    }

    [void]$out.Add("   smtp      : ${SmtpServer}:${Port}   UseSsl=$($UseSsl.IsPresent)")
    [void]$out.Add("   from      : $From")
    [void]$out.Add("   to        : $To")

    # The four things that actually go wrong, matched on the message text.
    $all = ($ex.Message + " " + $(if ($ex.InnerException) { $ex.InnerException.Message } else { "" }))
    switch -Regex ($all) {
        'could not be resolved|No such host|not known' {
            [void]$out.Add("   >> The SMTP server name does not resolve. Check smtp_server in")
            [void]$out.Add("      config.ini - the shipped default is a placeholder.")
            break
        }
        'relay|5\.7\.1|not permitted|denied' {
            [void]$out.Add("   >> The relay refused this sender. Ask your mail team to allow")
            [void]$out.Add("      this host's IP to relay, or use a from_address on an")
            [void]$out.Add("      accepted domain.")
            break
        }
        'authentic|5\.7\.0|must issue.*STARTTLS|credential' {
            [void]$out.Add("   >> The relay wants authentication or TLS. Try use_ssl = true")
            [void]$out.Add("      with smtp_port = 587. This script sends no password by")
            [void]$out.Add("      design, so the relay must accept anonymous submission.")
            break
        }
        'refused|timed out|unable to connect|actively refused' {
            [void]$out.Add("   >> Nothing is listening on that host and port. Test with:")
            [void]$out.Add("      Test-NetConnection $SmtpServer -Port $Port")
            break
        }
        default {
            [void]$out.Add("   >> Test the relay by hand with the same values:")
            [void]$out.Add("      Send-MailMessage -SmtpServer $SmtpServer -Port $Port ``")
            [void]$out.Add("        -From $From -To $To -Subject test -Body test")
        }
    }

    foreach ($line in $out) { [Console]::Error.WriteLine($line) }
    exit 1
}
