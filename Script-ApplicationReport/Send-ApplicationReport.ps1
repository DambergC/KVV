<#
.SYNOPSIS
  Sends per-application SCCM software installation summaries as HTML emails.

.DESCRIPTION
  - Reads config from Send-ApplicationReport.xml
  - Uses dbatools (Invoke-DbaQuery) to query the CM_ database
  - Builds HTML with PSWriteHTML
  - Sends HTML body via Send-MailKitMessage (Send-MailKitMessage module)
  - Sends one email per recipient

.SWITCHES
  -DryRun
      Test mode. No emails sent.
      HTML is generated and opened in the default browser.

  -MailOnly
      Only send HTML in email body.
      Do not write HTML files to disk.

  -AttachHTML
      Enable sending HTML files as attachments.
      Then, per-recipient, XML decides if they get the attachment (attach="true").
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$MailOnly,
    [switch]$AttachHTML
)

$scriptversion = '2.2'
$scriptname    = $MyInvocation.MyCommand.Name

Write-Host "Script: $scriptname  Version: $scriptversion"
if ($DryRun)     { Write-Host "Mode: DRY RUN (no emails will be sent, HTML will be opened)" }
if ($MailOnly)   { Write-Host "Mode: MAIL ONLY (no HTML files will be written)" }
if ($AttachHTML) { Write-Host "Mode: ATTACH HTML ENABLED (per-recipient attach flag from XML will be honored)" }

# ------------------------------------------------------------------------------------
# Load configuration XML
# ------------------------------------------------------------------------------------
[System.Xml.XmlDocument]$xml = Get-Content -Path (Join-Path $PSScriptRoot 'Send-ApplicationReport.xml')

$SQLserver      = $xml.Configuration.SQLServer
$SMTP           = $xml.Configuration.MailSMTP
$MailFrom       = $xml.Configuration.Mailfrom
$MailPortnumber = [int]$xml.Configuration.MailPort
$MailCustomer   = $xml.Configuration.MailCustomer
$HtmlPath       = $xml.Configuration.HTMLfilePath

# ------------------------------------------------------------------------------------
# Load required modules
# ------------------------------------------------------------------------------------
$requiredModules = @('Send-MailKitMessage', 'PSWriteHTML', 'dbatools')

foreach ($m in $requiredModules) {
    if (-not (Get-Module -Name $m -ListAvailable)) {
        Write-Host "Required module '$m' is not installed or not available in PSModulePath. Exiting."
        throw "Missing required module: $m"
    }

    Import-Module $m -ErrorAction Stop
}

# ------------------------------------------------------------------------------------
# SQL config
# ------------------------------------------------------------------------------------
$SqlServer = $SQLserver
$Database  = 'CM_KV1'

# Template SQL: [PRODUCT_FILTER] will be replaced for each Application from XML
$queryTemplate = @"
SELECT
    s.Publisher0      AS Publisher,
    s.ProductName0    AS ProductName,
    s.ProductVersion0 AS ProductVersion,
    COUNT(*)          AS InstallCount
FROM
    v_GS_INSTALLED_SOFTWARE AS s
    INNER JOIN System_DATA AS sys
        ON s.ResourceID = sys.MachineID
WHERE
   s.ProductName0 LIKE '[PRODUCT_FILTER]'
GROUP BY
    s.Publisher0,
    s.ProductName0,
    s.ProductVersion0
ORDER BY
    s.Publisher0,
    s.ProductName0,
    s.ProductVersion0;
"@

# ------------------------------------------------------------------------------------
# Process each <Application> from XML
# ------------------------------------------------------------------------------------
$applications = $xml.Configuration.Applications.Application
if (-not $applications) {
    Write-Warning "No <Applications><Application> entries found in Send-ApplicationReport.xml."
    return
}

foreach ($app in $applications) {
    $appName       = $app.Name
    $productFilter = $app.ProductFilter

    if ([string]::IsNullOrWhiteSpace($appName)) {
        Write-Warning "An <Application> without a Name attribute was found. Skipping."
        continue
    }

    if ([string]::IsNullOrWhiteSpace($productFilter)) {
        Write-Warning "Application '$appName' has no <ProductFilter> defined. Skipping."
        continue
    }

    Write-Host "=== Processing application '$appName' (filter: $productFilter) ==="

    # Build query for this application
    $query = $queryTemplate.Replace('[PRODUCT_FILTER]', $productFilter)

    # --------------------------------------------------------------------------------
    # Run query via dbatools
    # --------------------------------------------------------------------------------
    try {
        $dt = Invoke-DbaQuery -SqlInstance $SqlServer -Database $Database -Query $query -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to run query for application '$appName': $($_.Exception.Message)"
        continue
    }

    # Ensure $dt is an array (0,1, or many rows)
    $dt = @($dt)

    # Remove DataRow internal properties; keep only columns we actually use
    if ($dt.Count -gt 0) {
        $dt = $dt | Select-Object Publisher, ProductName, ProductVersion, InstallCount
    }

    # --------------------------------------------------------------------------------
    # Build HTML with PSWriteHTML (returned as a string)
    # --------------------------------------------------------------------------------
    $reportTitle = "$appName Software Installation Summary"
    $now         = Get-Date -Format "yyyy-MM-dd HH:mm"

    $html = New-HTML -TitleText $reportTitle -Online {
        # CSS
        New-HTMLTag -Tag 'style' {
            @"
table {
    border-collapse: collapse;
}
table, th, td {
    border: 1px solid #cccccc;
}
th, td {
    padding: 4px 8px;
    font-size: 11px;
}
.header-block {
    margin: 10px auto;
    font-size: 12px;
    text-align: center;
    max-width: 600px;
}
.header-block div {
    margin: 2px 0;
}
.header-label {
    font-weight: bold;
}
"@
        }

        # Section 1: centered header + summary info
        New-HTMLSection -HeaderText $reportTitle {

            New-HTMLTag -Tag 'div' -Attributes @{ class = 'header-block' } {
                New-HTMLTag -Tag 'div' {
                    "<span class='header-label'>Report generated:</span> $($now)"
                }
                New-HTMLTag -Tag 'div' {
                    "<span class='header-label'>Customer:</span> $MailCustomer"
                }
                New-HTMLTag -Tag 'div' {
                    "<span class='header-label'>Filter:</span> ProductName LIKE '$productFilter'"
                }

                if ($dt.Count -gt 0) {
                    $totalInstalls = ($dt | Measure-Object -Property InstallCount -Sum).Sum
                    New-HTMLTag -Tag 'div' {
                        "<span class='header-label'>Total installations (sum of InstallCount):</span> $totalInstalls"
                    }
                }
                else {
                    New-HTMLTag -Tag 'div' {
                        "<span class='header-label'>Total installations (sum of InstallCount):</span> 0"
                    }
                }
            }
        }

        # Section 2: table only (or "no data" message)
        New-HTMLSection {
            if ($dt.Count -eq 0) {
                New-HTMLText -Text "No matching installations found." -Color Red -FontSize 14
            }
            else {
                New-HTMLTable -DataTable $dt
            }
        }
    }

    # --------------------------------------------------------------------------------
    # Handle HTML file output (skipped when -MailOnly is used)
    # --------------------------------------------------------------------------------
    $htmlFile = $null
    if (-not $MailOnly -and -not [string]::IsNullOrWhiteSpace($HtmlPath)) {
        if (-not (Test-Path -LiteralPath $HtmlPath)) {
            New-Item -ItemType Directory -Path $HtmlPath -Force | Out-Null
        }
        $safeAppName = ($appName -replace '[^a-zA-Z0-9_-]', '_')
        $htmlFile    = Join-Path $HtmlPath ("{0}_{1}.html" -f $safeAppName, (Get-Date -Format 'yyyyMMdd_HHmm'))
        $html | Out-File -FilePath $htmlFile -Encoding UTF8
        Write-Host "Saved HTML report for '$appName' to $htmlFile"
    }

    # If DryRun: open HTML and skip sending email
    if ($DryRun) {
        Write-Host "DryRun: opening HTML for '$appName' and skipping email."
        if ($htmlFile) {
            Start-Process $htmlFile
        }
        else {
            # If MailOnly + DryRun (no file) – create temp file to view
            $tempFile = [System.IO.Path]::GetTempFileName().Replace('.tmp','.html')
            $html | Out-File -FilePath $tempFile -Encoding UTF8
            Start-Process $tempFile
        }
        continue
    }

    # --------------------------------------------------------------------------------
    # Build recipient list from XML (with per-recipient attach flag)
    # --------------------------------------------------------------------------------
    $recipientNodes = $app.Recipients.Recipient
    if (-not $recipientNodes) {
        Write-Warning "Application '$appName' has no <Recipients><Recipient> entries. Skipping email."
        continue
    }

    # Build an object per recipient: Email + AttachHtmlFromXml (bool)
    $recipients = @()
    foreach ($r in $recipientNodes) {
        if (-not $r.email) { continue }

        $attachAttr = $r.attach
        $attachFlag = $false
        if ($attachAttr -and $attachAttr.ToString().ToLower() -eq 'true') {
            $attachFlag = $true
        }

        $recipients += [pscustomobject]@{
            Email        = $r.email
            AttachFromXml = $attachFlag
        }
    }

    if ($recipients.Count -eq 0) {
        Write-Warning "Application '$appName' has Recipient nodes but no 'email' values. Skipping email."
        continue
    }

    Write-Host "Recipients for '$appName': $($recipients.Email -join ', ')"

    # --------------------------------------------------------------------------------
    # Prepare From address (MailboxAddress reused for all recipients)
    # --------------------------------------------------------------------------------
    $fromAddress = New-Object MimeKit.MailboxAddress ('', $MailFrom)
    $Subject     = "$appName Software Installation Summary - $now"

    # Base attachment path (if file exists)
    $baseAttachment = $null
    if ($htmlFile -and (Test-Path -LiteralPath $htmlFile)) {
        $baseAttachment = $htmlFile
    }

    # --------------------------------------------------------------------------------
    # Send one email per recipient
    # --------------------------------------------------------------------------------
    foreach ($rec in $recipients) {
        $addr          = $rec.Email
        $attachFromXml = $rec.AttachFromXml

        if ([string]::IsNullOrWhiteSpace($addr)) { continue }

        Write-Host "Sending mail for '$appName' to $addr (attach from XML: $attachFromXml) ..."

        # RecipientList expects a MimeKit.InternetAddress (MailboxAddress derives from it)
        $recipientAddress = New-Object MimeKit.MailboxAddress ('', $addr)

        $params = @{
            SMTPServer    = $SMTP
            Port          = $MailPortnumber
            From          = $fromAddress
            RecipientList = $recipientAddress
            Subject       = $Subject
            HTMLBody      = $html
        }

        # Only attach if:
        # - global switch -AttachHTML is set
        # - a base attachment file exists
        # - this recipient's XML has attach="true"
        if ($AttachHTML -and $baseAttachment -and $attachFromXml) {
            $params['AttachmentList'] = @($baseAttachment)
            Write-Host " -> Attaching HTML for recipient $addr"
        }
        else {
            Write-Host " -> No attachment for recipient $addr"
        }

        Send-MailKitMessage @params

        Write-Host "Mail for '$appName' sent to $addr."
    }

    Write-Host ""
}

Write-Host "All applications processed."
