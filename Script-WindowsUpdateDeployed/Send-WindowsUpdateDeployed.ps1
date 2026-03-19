<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Updatestatus for Windows update thru MECM

.DESCRIPTION
   Creates an HTML report for one Software Update Group specified in an XML file and sends an email.
   Intended to run as a scheduled task on the site server.

   Requires:
   - PSWriteHTML (for New-HTML)
   - Send-MailKitMessage
   - (Your MECM modules/logic to collect update deployment data)

   Expected XML file: Send-WindowsUpdateDeployedv2.xml (next to script)
-------------------------------------------------------------------------------------------------------------------------
#>

$ErrorActionPreference = 'Stop'

$scriptname = $MyInvocation.MyCommand.Name
$today      = (Get-Date).ToString('yyyy-MM-dd')

#########################################################
# Load configuration XML
#########################################################
$xmlPath = Join-Path -Path $PSScriptRoot -ChildPath 'Send-WindowsUpdateDeployedv2.xml'
if (-not (Test-Path -LiteralPath $xmlPath)) {
    throw "Missing configuration XML: $xmlPath"
}
[xml]$xml = Get-Content -LiteralPath $xmlPath

#########################################################
# XML helpers (PS 5.1 compatible)
#########################################################
function Get-XmlText {
    param(
        [Parameter(Mandatory)][object]$Node,
        [Parameter(Mandatory)][string]$Path
    )
    $n = $Node.SelectSingleNode($Path)
    if (-not $n) { return $null }

    $t = $n.InnerText
    if ($null -eq $t) { return $null }

    $t = $t.Trim()
    if ($t -eq '') { return $null }
    return $t
}

function Get-XmlInt {
    param(
        [Parameter(Mandatory)][object]$Node,
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][int]$Default = 0
    )
    $t = Get-XmlText -Node $Node -Path $Path
    if ($null -eq $t) { return $Default }

    $i = 0
    if ([int]::TryParse($t, [ref]$i)) { return $i }
    return $Default
}

function Get-XmlIntArray {
    param(
        [Parameter(Mandatory)][object]$Node,
        [Parameter(Mandatory)][string]$Path
    )

    $nodes = $Node.SelectNodes($Path)
    if (-not $nodes) { return @() }

    $values = @()
    foreach ($n in @($nodes)) {
        $t = $n.InnerText
        if ($null -eq $t) { continue }
        $t = $t.Trim()
        if ($t -eq '') { continue }

        $i = 0
        if ([int]::TryParse($t, [ref]$i)) { $values += $i }
    }

    return @($values)
}

#########################################################
# Read configuration values
#########################################################
# Logging (optional)
$Logfilepath      = Get-XmlText -Node $xml -Path '/Configuration/Logfile/Path'
$logfilename      = Get-XmlText -Node $xml -Path '/Configuration/Logfile/Name'
$Logfilethreshold = Get-XmlInt  -Node $xml -Path '/Configuration/Logfile/Logfilethreshold' -Default 0

$Logfile = $null
if ($Logfilepath -and $logfilename) {
    $Logfile = Join-Path -Path $Logfilepath -ChildPath $logfilename
}

# General / schedule
$SiteServer = Get-XmlText -Node $xml -Path '/Configuration/SiteServer'

$LimitDays                    = Get-XmlInt  -Node $xml -Path '/Configuration/UpdateDeployed/LimitDays' -Default 0
$DaysAfterPatchTuesdayToReport = Get-XmlInt  -Node $xml -Path '/Configuration/UpdateDeployed/DaysAfterPatchToRun' -Default 0
$UpdateGroupName              = Get-XmlText -Node $xml -Path '/Configuration/UpdateDeployed/UpdateGroupName'

# Mail
$SMTP           = Get-XmlText -Node $xml -Path '/Configuration/MailSMTP'
$MailFrom       = Get-XmlText -Node $xml -Path '/Configuration/Mailfrom'
$MailPortnumber = Get-XmlInt  -Node $xml -Path '/Configuration/MailPort' -Default 25
$MailCustomer   = Get-XmlText -Node $xml -Path '/Configuration/MailCustomer'

#########################################################
# Recipients (support both schemas; PS 5.1 compatible)
#########################################################
$MailTo = @()

# A: /Configuration/Recipients/Recipients/Email
$recipientEmailNodes = $xml.SelectNodes('/Configuration/Recipients/Recipients/Email')
foreach ($n in @($recipientEmailNodes)) {
    $email = $null
    if ($n -and $n.InnerText) { $email = $n.InnerText.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($email)) { $MailTo += $email }
}

# B: /Configuration/Recipients/Recipients[@email]
$recipientAttrNodes = $xml.SelectNodes('/Configuration/Recipients/Recipients[@email]')
foreach ($n in @($recipientAttrNodes)) {
    $email = $null
    if ($n) { $email = $n.GetAttribute('email') }
    if ($email) { $email = $email.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($email)) { $MailTo += $email }
}

# unique
$MailTo = $MailTo | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

#########################################################
# Disable report months (support both schemas; PS 5.1 compatible)
#########################################################
$DisableReportMonths = @()
$DisableReportMonths += Get-XmlIntArray -Node $xml -Path '/Configuration/DisableReportMonth/DisableReportMonth/Number'

$monthAttrNodes = $xml.SelectNodes('/Configuration/DisableReportMonth/DisableReportMonth[@Number]')
foreach ($n in @($monthAttrNodes)) {
    $t = $null
    if ($n) { $t = $n.GetAttribute('Number') }
    if ($t) { $t = $t.Trim() }
    $i = 0
    if (-not [string]::IsNullOrWhiteSpace($t) -and [int]::TryParse($t, [ref]$i)) {
        $DisableReportMonths += $i
    }
}
$DisableReportMonths = $DisableReportMonths | Select-Object -Unique

#########################################################
# Basic validation
#########################################################
foreach ($required in @(
    @{ Name = 'SiteServer'; Value = $SiteServer },
    @{ Name = 'SMTP'; Value = $SMTP },
    @{ Name = 'MailFrom'; Value = $MailFrom },
    @{ Name = 'MailCustomer'; Value = $MailCustomer },
    @{ Name = 'UpdateGroupName'; Value = $UpdateGroupName }
)) {
    if ([string]::IsNullOrWhiteSpace($required.Value)) {
        throw "Missing required XML configuration value: $($required.Name)"
    }
}
if ($MailTo.Count -eq 0) {
    throw "No recipients found in XML. Supported: /Configuration/Recipients/Recipients/Email and /Configuration/Recipients/Recipients[@email]"
}

#########################################################
# Logging helpers (optional)
#########################################################
function Write-Log {
    param([Parameter(Mandatory)][string]$LogString)

    if ([string]::IsNullOrWhiteSpace($Logfile)) { return }

    try {
        $stamp = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        Add-Content -Path $Logfile -Value "$stamp $LogString"
    } catch {
        Write-Warning "Failed to write to log file '$Logfile': $($_.Exception.Message)"
    }
}

Write-Log -LogString "======================= $scriptname - Script START ============================="

#########################################################
# Module imports
#########################################################
function Ensure-Module {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][ScriptBlock]$ImportScript
    )

    if (-not (Get-Module -Name $Name -ListAvailable)) {
        throw "Required module '$Name' is not available on this machine."
    }

    if (-not (Get-Module -Name $Name)) {
        if ($ImportScript) { & $ImportScript } else { Import-Module -Name $Name }
    }
}

Ensure-Module -Name 'PSWriteHTML' -ImportScript { Import-Module PSWriteHTML }
Ensure-Module -Name 'Send-MailKitMessage' -ImportScript { Import-Module Send-MailKitMessage }

#########################################################
# Time context (month/year used in title)
#########################################################
$monthname = (Get-Date).ToString('MMMM')
$year      = (Get-Date).ToString('yyyy')

# Optional: skip if current month is disabled
$currentMonthNumber = [int](Get-Date).ToString('MM')
if ($DisableReportMonths -contains $currentMonthNumber) {
    Write-Log -LogString "Month $currentMonthNumber is disabled in config. Exiting."
    return
}



#########################################################
# Function to collect data about updates
#########################################################
function Get-UpdatesDeployedThisMonth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UpdateGroupName,
        [Parameter()][int]$Month = (Get-Date).Month,
        [Parameter()][int]$Year  = (Get-Date).Year,
        [Parameter()][datetime]$SinceDate
    )

    if (-not (Get-Command -Name Get-CMSoftwareUpdate -ErrorAction SilentlyContinue)) {
        throw "Get-CMSoftwareUpdate is not available. Ensure the ConfigurationManager module is imported."
    }

    # Auto-detect the CM PSDrive (PSProvider: CMSite)
    $cmDrive = Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $cmDrive) {
        throw "No Configuration Manager PSDrive found (PSProvider 'CMSite'). Import ConfigurationManager and connect to a site (e.g. Set-Location '<SiteCode>:')."
    }

    $cmPath = "$($cmDrive.Name):"

    Push-Location
    try {
        Set-Location -Path $cmPath

        $updates = @(Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName)

        $filtered = if ($PSBoundParameters.ContainsKey('SinceDate') -and $SinceDate) {
            @($updates | Where-Object { $_.DatePosted -and ($_.DatePosted -ge $SinceDate) })
        }
        else {
            @($updates | Where-Object { $_.DatePosted -and ($_.DatePosted.Month -eq $Month) -and ($_.DatePosted.Year -eq $Year) })
        }

        $rows = foreach ($u in ($filtered | Sort-Object DatePosted -Descending)) {
            [pscustomobject]@{
                ArticleID            = $u.ArticleID
                Title                = $u.LocalizedDisplayName
                LocalizedDescription = $u.LocalizedDescription
                DatePosted           = $u.DatePosted
                Deployed             = $u.IsDeployed
                URL                  = $u.LocalizedInformativeURL
                Severity             = $u.SeverityName
            }
        }

        return @($rows)
    }
    finally {
        Pop-Location
    }
}
#########################################################
# Collect Windows Update deployment data
#########################################################
$updateRows = Get-UpdatesDeployedThisMonth -UpdateGroupName $UpdateGroupName
$updateCount = @($updateRows).Count

#########################################################
# Build HTML report (PSWriteHTML) - SAME PATTERN AS DPSTATUS
#########################################################
$reportTitle = "$MailCustomer - Windows Updates $monthname $year"

# NOTE: This is the SAME base64 image string used in Send-CheckDPStatus.ps1.
# If you want a different logo/image, replace it here.
$base64Img = @"
data:image/png;base64,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
"@

# The base64 string is extremely long; keep it exactly as-is from DPStatus in your repo.
# If you paste the whole DPStatus img tag line, replace $base64Img with that full string.

$html = New-HTML -TitleText $reportTitle -Online {
    New-HTMLTag -Tag 'style' {
@"
body { font-family: Arial, Helvetica, sans-serif; }
.small { font-size: 11px; color: #333; }
.summary { margin: 10px 0 20px 0; }
.summary div { margin: 2px 0; }
"@
    }

    # Image (same idea as DPStatus)
    New-HTMLTag -Tag 'img' -Attributes @{
        alt = 'My image'
        src = $base64Img
    }

    # Summary FIRST (same style as DPStatus, but WindowsUpdate values)
    New-HTMLSection -HeaderText "Summary" -HeaderTextSize 20 -HeaderBackGroundColor Darkblue {
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'summary' } {
            New-HTMLTag -Tag 'div' { "Datum: <b>$today</b>" }
            New-HTMLTag -Tag 'div' { "SiteServer: <b>$SiteServer</b>" }
            New-HTMLTag -Tag 'div' { "UpdateGroupName: <b>$UpdateGroupName</b>" }
            New-HTMLTag -Tag 'div' { "Antal uppdateringar i rapport: <b>$updateCount</b>" }
            New-HTMLTag -Tag 'div' { "LimitDays: <b>$LimitDays</b>" }
            New-HTMLTag -Tag 'div' { "DaysAfterPatchTuesdayToReport: <b>$DaysAfterPatchTuesdayToReport</b>" }
        }
    }

    # Main section: Updates table
    New-HTMLSection -HeaderText "Windows Updates ($updateCount)" -HeaderTextSize 20 -HeaderBackGroundColor Darkblue {
        if ($updateCount -eq 0) {
            New-HTMLText -Text 'Inga poster.' -Color Gray
        } else {
            # Select only columns that exist in your dataset
            New-HTMLTable -DataTable ($updateRows | Select-Object Title, LocalizedDescription, ArticleID, DatePosted, Deployed, URL, Severity)
        }
    }

    # Footer (same as DPStatus)
    New-HTMLTag -Tag 'p' -Attributes @{ class = 'small' } {
        "Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b>"
    }
}

#########################################################
# Send email (Send-MailKitMessage) - SAME STYLE AS DPSTATUS
#########################################################
# Build recipient list (MimeKit)
Add-Type -AssemblyName 'System.Runtime' -ErrorAction SilentlyContinue | Out-Null

$RecipientList = [MimeKit.InternetAddressList]::new()
foreach ($addr in $MailTo) {
    if ([string]::IsNullOrWhiteSpace($addr)) { continue }
    $RecipientList.Add([MimeKit.InternetAddress]$addr) | Out-Null
}

$Subject = [string]"$MailCustomer - Windows Updates $monthname $year - $today"

$Parameters = @{
    SMTPServer                    = $SMTP
    Port                          = $MailPortnumber
    From                          = [MimeKit.MailboxAddress]$MailFrom
    RecipientList                 = $RecipientList
    Subject                       = $Subject
    HTMLBody                      = $html
    AttachmentList                = [System.Collections.Generic.List[string]]::new()
}

try {
    Send-MailKitMessage @Parameters
    Write-Log -LogString "$scriptname - Mail sent (or queued by SMTP) to: $($MailTo -join ', ')"
} catch {
    Write-Host "ERROR: Failed to send mail. $($_.Exception.Message)" -ForegroundColor Red
    Write-Log -LogString "ERROR: Failed to send mail. $($_.Exception.Message)"
}

Write-Log -LogString "$scriptname - Script end"
