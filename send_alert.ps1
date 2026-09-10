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
    Write-Error ("MAIL FAILED: " + $_.Exception.Message)
    exit 1
}
