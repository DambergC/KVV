<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Updatestatus for Windows update thru MECM

.DESCRIPTION
   Creates an HTML report for one or more Software Update deployments specified in an XML file and sends an email.
   Intended to run as a scheduled task on the site server.

   This version follows the same structure as Script-DPstatus/Send-CheckDPStatus.ps1:
   - Configured via XML (no parameters)
   - Robust XML helpers + validation
   - Module import helpers
   - Consistent logging

   Requires:
   - ConfigurationManager (ConfigMgr console installed, SMS_ADMIN_UI_PATH set)
   - Send-MailKitMessage

   Optional:
   - PSWriteHTML (not required; this script uses ConvertTo-Html like the original)

   Expected XML file: ScriptConfigPatchReleased.xml (next to script)

.NOTES
   - Scripts are offered AS IS with no warranty.
   - Test in non-production first.
-------------------------------------------------------------------------------------------------------------------------
#>

$ErrorActionPreference = 'Stop'

$scriptname = $MyInvocation.MyCommand.Name
$today      = (Get-Date).ToString('yyyy-MM-dd')

#########################################################
# Load configuration XML
#########################################################
$xmlPath = Join-Path -Path $PSScriptRoot -ChildPath 'ScriptConfigPatchReleased.xml'
if (-not (Test-Path -LiteralPath $xmlPath)) {
    throw "Missing configuration XML: $xmlPath"
}

[xml]$xml = Get-Content -LiteralPath $xmlPath

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

function Get-XmlBool {
    param(
        [Parameter(Mandatory)][object]$Node,
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][bool]$Default = $false
    )
    $t = Get-XmlText -Node $Node -Path $Path
    if ($null -eq $t) { return $Default }
    try { return [System.Convert]::ToBoolean($t) } catch { return $Default }
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

    $values = foreach ($n in @($nodes)) {
        $t = $n.InnerText
        if ($null -eq $t) { continue }
        $t = $t.Trim()
        if ($t -eq '') { continue }
        $i = 0
        if ([int]::TryParse($t, [ref]$i)) { $i }
    }

    return @($values)
}

#########################################################
# Read configuration values (tolerant of old/new XML schema)
#########################################################
# Logging
$Logfilepath      = Get-XmlText -Node $xml -Path '/Configuration/Logfile/Path'
$logfilename      = Get-XmlText -Node $xml -Path '/Configuration/Logfile/Name'
$Logfilethreshold = Get-XmlInt  -Node $xml -Path '/Configuration/Logfile/Logfilethreshold' -Default 0

$Logfile = $null
if ($Logfilepath -and $logfilename) {
    $Logfile = Join-Path -Path $Logfilepath -ChildPath $logfilename
}

# General / schedule
$SiteServer = Get-XmlText -Node $xml -Path '/Configuration/SiteServer'

$LimitDays                   = Get-XmlInt  -Node $xml -Path '/Configuration/UpdateDeployed/LimitDays' -Default 0
$DaysAfterPatchTuesdayToReport = Get-XmlInt -Node $xml -Path '/Configuration/UpdateDeployed/DaysAfterPatchToRun' -Default 0
$UpdateGroupName             = Get-XmlText -Node $xml -Path '/Configuration/UpdateDeployed/UpdateGroupName'

# Mail
$SMTP           = Get-XmlText -Node $xml -Path '/Configuration/MailSMTP'
$MailFrom       = Get-XmlText -Node $xml -Path '/Configuration/Mailfrom'
$MailPortnumber = Get-XmlInt  -Node $xml -Path '/Configuration/MailPort' -Default 25
$MailCustomer   = Get-XmlText -Node $xml -Path '/Configuration/MailCustomer'

# Recipients (support both the original structure and a more explicit one)
$recipientNodes = $xml.SelectNodes('/Configuration/Recipients/Recipients/Email')
$MailTo = @()
foreach ($r in @($recipientNodes)) {
    $email = $null
    if ($r -and $r.InnerText) {
        $email = $r.InnerText.Trim()
    }
    if ($email) { $MailTo += $email }
}
$MailTo = $MailTo | Select-Object -Unique

# Disable report months
$DisableReportMonths = Get-XmlIntArray -Node $xml -Path '/Configuration/DisableReportMonth/DisableReportMonth/Number'

#########################################################
# Basic validation (log file is optional)
#########################################################
foreach ($required in @(
    @{ Name = 'SiteServer'; Value = $SiteServer },
    @{ Name = 'SMTP'; Value = $SMTP },
    @{ Name = 'MailFrom'; Value = $MailFrom },
    @{ Name = 'UpdateGroupName'; Value = $UpdateGroupName }
)) {
    if ([string]::IsNullOrWhiteSpace($required.Value)) {
        throw "Missing required XML configuration value: $($required.Name)"
    }
}

if ($MailTo.Count -eq 0) {
    throw "No recipients found in XML at /Configuration/Recipients/Recipients/Email"
}

Write-Host "Script: $scriptname" -ForegroundColor Cyan
Write-Host "SiteServer: $SiteServer" -ForegroundColor Cyan
Write-Host "UpdateGroupName: $UpdateGroupName" -ForegroundColor Cyan

#########################################################
# Logging
#########################################################
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$LogString
    )

    if ([string]::IsNullOrWhiteSpace($Logfile)) {
        return
    }

    try {
        $stamp = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        Add-Content -Path $Logfile -Value "$stamp $LogString"
    } catch {
        Write-Warning "Failed to write to log file '$Logfile': $($_.Exception.Message)"
    }
}

function Rotate-Log {
    if ([string]::IsNullOrWhiteSpace($Logfilepath) -or $Logfilethreshold -le 0) {
        return
    }

    try {
        $target = Get-ChildItem -LiteralPath $Logfilepath -Filter "windo*.log" -ErrorAction Stop
    } catch {
        Write-Warning "Rotate-Log: Unable to enumerate log files in '$Logfilepath': $($_.Exception.Message)"
        return
    }

    $datetime = Get-Date -uformat "%Y-%m-%d-%H%M"

    foreach ($file in $target) {
        try {
            if ($file.Length -ge $Logfilethreshold) {
                Write-Host "file named $($file.Name) is bigger than $Logfilethreshold B"
                $newname = "$($file.BaseName)_${datetime}.log"
                Rename-Item -LiteralPath $file.FullName -NewName $newname -ErrorAction Stop

                $oldLogDir = Join-Path -Path $Logfilepath -ChildPath 'OLDLOG'
                if (-not (Test-Path -LiteralPath $oldLogDir)) {
                    New-Item -Path $oldLogDir -ItemType Directory -Force | Out-Null
                }

                $movedSource = Join-Path -Path $Logfilepath -ChildPath $newname
                Move-Item -LiteralPath $movedSource -Destination $oldLogDir -Force
                Write-Host "Done rotating file"
            } else {
                Write-Host "file named $($file.Name) is not bigger than $Logfilethreshold B"
            }
        } catch {
            Write-Warning "Rotate-Log: Failed processing '$($file.FullName)': $($_.Exception.Message)"
        }
    }

    Write-Host "Logfile checked!"
}

Write-Log -LogString "======================= $scriptname - Script START ============================="
Rotate-Log

#########################################################
# SiteCode helper (WMI/DCOM)
#########################################################
function Get-CMSiteCode {
    param(
        [Parameter(Mandatory)][string]$ComputerName
    )

    try {
        $siteCode = Get-WmiObject -Namespace 'root\SMS' -Class SMS_ProviderLocation -ComputerName $ComputerName |
            Select-Object -First 1 -ExpandProperty SiteCode
    } catch {
        throw "Unable to determine SiteCode via WMI from SMS_ProviderLocation on '$ComputerName'. Error: $($_.Exception.Message)"
    }

    if (-not $siteCode) {
        throw "Unable to determine SiteCode via WMI from SMS_ProviderLocation on '$ComputerName' (no SiteCode returned)."
    }

    return $siteCode
}

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

# ConfigurationManager import
if (-not (Get-Module -Name ConfigurationManager)) {
    if (-not $Env:SMS_ADMIN_UI_PATH) {
        throw "SMS_ADMIN_UI_PATH is not set. Run on a machine with the ConfigMgr console installed."
    }

    $cmModulePath = Join-Path -Path ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5)) -ChildPath 'ConfigurationManager.psd1'
    Import-Module $cmModulePath
}

Ensure-Module -Name 'Send-MailKitMessage' -ImportScript { Import-Module Send-MailKitMessage }

# Optional modules used elsewhere in the repo
if (Get-Module -Name PSWriteHTML -ListAvailable) {
    Import-Module PSWriteHTML -ErrorAction SilentlyContinue
}
if (Get-Module -Name PatchManagementSupportTools -ListAvailable) {
    Import-Module PatchManagementSupportTools -ErrorAction SilentlyContinue
}

if (Get-Command -Name Get-CMModule -ErrorAction SilentlyContinue) {
    Get-CMModule | Out-Null
}

#########################################################
# Connect to CM PSDrive
#########################################################
$sitecode = Get-CMSiteCode -ComputerName $SiteServer
$setSiteCode = "$sitecode`:"
Write-Log -LogString "$scriptname - SiteCode resolved: $sitecode"

Push-Location
try {
    Set-Location $setSiteCode
} catch {
    Pop-Location
    throw "Failed to Set-Location to ConfigMgr PSDrive '$setSiteCode'. Error: $($_.Exception.Message)"
}

try {
    #########################################################
    # Section to extract monthname, year and weeknumbers
    #########################################################
    function Get-ISO8601Week {
        # Adapted from https://stackoverflow.com/a/43736741/444172
        [CmdletBinding()]
        param(
            [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
            [datetime]$DateTime
        )
        process {
            foreach ($_DateTime in $DateTime) {
                $resultObject = [pscustomobject]@{
                    Year       = $null
                    WeekNumber = $null
                    WeekString = $null
                    DateString = $_DateTime.ToString('yyyy-MM-dd   dddd')
                }

                $dayOfWeek = $_DateTime.DayOfWeek.value__
                if ($dayOfWeek -eq 0) { $dayOfWeek = 7 }

                $_DateTime = $_DateTime.AddDays((4 - $dayOfWeek))

                $resultObject.WeekNumber = [math]::Ceiling($_DateTime.DayOfYear / 7)
                $resultObject.Year       = $_DateTime.Year
                $resultObject.WeekString = "$($_DateTime.Year)-W$("$($resultObject.WeekNumber)".PadLeft(2, '0'))"

                Write-Output $resultObject
            }
        }
    }

    $monthname = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month)
    $year      = (Get-Date).Year

    $todayDefault = Get-Date
    $todayshort   = $todayDefault.ToShortDateString()

    # Patch Tuesday (requires PatchManagementSupportTools, but keep behavior compatible)
    if (-not (Get-Command -Name Get-PatchTuesday -ErrorAction SilentlyContinue)) {
        throw "Get-PatchTuesday was not found. Install/import PatchManagementSupportTools or adjust the script."
    }

    $thismonth = $todayDefault.Month
    $patchTuesdayThisMonth = Get-PatchTuesday -Month $thismonth -Year $todayDefault.Year

    $ReportdayCompare = ($patchTuesdayThisMonth.AddDays($DaysAfterPatchTuesdayToReport)).ToString('yyyy-MM-dd')

    if ($DisableReportMonths -and ($todayDefault.Month -in $DisableReportMonths)) {
        Write-Log -LogString "$scriptname - This month is skipped"
        Write-Log -LogString "$scriptname - Script exit!"
        return
    }

    if ($todayshort -ne $ReportdayCompare) {
        Write-Log -LogString "========= $scriptname - Date not equal patchtuesday $($patchTuesdayThisMonth.ToShortDateString()) and its now $todayshort. This report will run $ReportdayCompare ========"
        Write-Log -LogString "========================== $scriptname - Script exit! =========================="
        return
    }

    #########################################################
    # Gather updates
    #########################################################
    Write-Log -LogString "=====================Processing Software Update Group $UpdateGroupName=========================="

    $result = @()
    $updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName

    foreach ($item in $updates) {
        $result += [pscustomobject]@{
            ArticleID            = $item.ArticleID
            Title                = $item.LocalizedDisplayName
            LocalizedDescription = $item.LocalizedDescription
            DatePosted           = $item.DatePosted
            Deployed             = $item.IsDeployed
            URL                  = $item.LocalizedInformativeURL
            Severity             = $item.SeverityName
        }
    }

    $limit = (Get-Date).AddDays($LimitDays)
    $UpdatesFound = $result | Sort-Object DatePosted -Descending | Where-Object { $_.DatePosted -ge $limit }

    #########################################################
    # Build HTML report (ConvertTo-Html as original)
    #########################################################
    $header = @"
<style>

    th {

        font-family: Arial, Helvetica, sans-serif;
        color: White;
        font-size: 12px;
        border: 1px solid black;
        padding: 3px;
        background-color: Black;

    }
    p {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 12px;

    }
    ol {

        font-family: Arial, Helvetica, sans-serif;
        list-style-type: square;
        color: black;
        font-size: 12px;

    }
    tr {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 11px;
        vertical-align: text-top;

    }

    body {
        background-color: lightgray;
      }
      table {
        border: 1px solid black;
        border-collapse: collapse;
      }

      td {
        border: 1px solid black;
        padding: 5px;
        background-color: #E0F3F7;
      }

</style>
"@

    # Keep the existing embedded logo from the original script by leaving it in XML or externalizing.
    # (If the logo is needed inline, add it back here.)
    $pre = @"
<p><h1>Windows updates $monthname $year</h1><br>
<p>The following updates from Microsoft are deployed this month.</p>
"@

    $post = @"
<p>Raport created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
"@

    $UpdatesFoundHtml = $UpdatesFound | ConvertTo-Html -Title "Downloaded patches" -PreContent $pre -PostContent $post -Head $header

    #########################################################
    # Mailsettings (Send-MailKitMessage)
    #########################################################
    $UseSecureConnectionIfAvailable = $false

    $SMTPServer = $SMTP
    $Port       = $MailPortnumber
    $From       = [MimeKit.MailboxAddress]$MailFrom

    $RecipientList = [MimeKit.InternetAddressList]::new()
    foreach ($Recipient in $MailTo) {
        $RecipientList.Add([MimeKit.InternetAddress]$Recipient)
    }

    $Subject  = [string]"WindowsUpdate $MailCustomer $monthname $year"
    $HTMLBody = [string]$UpdatesFoundHtml

    $AttachmentList = [System.Collections.Generic.List[string]]::new()

    $Parameters = @{ 
        UseSecureConnectionIfAvailable = $UseSecureConnectionIfAvailable
        SMTPServer                    = $SMTPServer
        Port                          = $Port
        From                          = $From
        RecipientList                 = $RecipientList
        Subject                       = $Subject
        HTMLBody                      = $HTMLBody
        AttachmentList                = $AttachmentList
    }

    Send-MailKitMessage @Parameters

    Write-Log -LogString "========================== $scriptname - Mail on it´s way to $RecipientList "
    Write-Log -LogString "========================== $scriptname - Script exit! =========================="
} finally {
    Pop-Location
}