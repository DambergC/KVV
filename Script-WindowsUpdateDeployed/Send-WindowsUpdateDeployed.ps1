<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Updatestatus for Windows update thru MECM

.DESCRIPTION
   Creates an HTML report for one Software Update Group specified in an XML file and sends an email.
   Intended to run as a scheduled task on the site server.

   Requires:
   - ConfigurationManager (ConfigMgr console installed, SMS_ADMIN_UI_PATH set)
   - Send-MailKitMessage
   - PatchManagementSupportTools (for Get-PatchTuesday)
   - PSWriteHTML (for New-HTML)

   Expected XML file: Send-WindowsUpdateDeployedv2.xml (next to script)
-------------------------------------------------------------------------------------------------------------------------
#>

$ErrorActionPreference = 'Stop'

$scriptname = $MyInvocation.MyCommand.Name

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
        if ([int]::TryParse($t, [ref]$i)) {
            $values += $i
        }
    }

    return @($values)
}

function Add-UniqueString {
    param(
        [Parameter(Mandatory)][System.Collections.ArrayList]$List,
        [Parameter(Mandatory)][string]$Value
    )
    # case-insensitive unique add for PS5.1
    foreach ($existing in $List) {
        if ($existing -and ($existing.ToString().ToLowerInvariant() -eq $Value.ToLowerInvariant())) {
            return
        }
    }
    [void]$List.Add($Value)
}

#########################################################
# Read configuration values
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

$LimitDays                     = Get-XmlInt  -Node $xml -Path '/Configuration/UpdateDeployed/LimitDays' -Default 0
$DaysAfterPatchTuesdayToReport  = Get-XmlInt  -Node $xml -Path '/Configuration/UpdateDeployed/DaysAfterPatchToRun' -Default 0
$UpdateGroupName               = Get-XmlText -Node $xml -Path '/Configuration/UpdateDeployed/UpdateGroupName'

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

# A: .../Number
$DisableReportMonths += Get-XmlIntArray -Node $xml -Path '/Configuration/DisableReportMonth/DisableReportMonth/Number'

# B: attribute Number="x"
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
    throw "No recipients found in XML. Supported: /Configuration/Recipients/Recipients/Email and /Configuration/Recipients/Recipients[@email]"
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

# PSWriteHTML for New-HTML (same approach as Send-ApplicationReport.ps1 uses)
Ensure-Module -Name 'PSWriteHTML' -ImportScript { Import-Module PSWriteHTML }

# PatchManagementSupportTools provides Get-PatchTuesday
if (-not (Get-Command -Name Get-PatchTuesday -ErrorAction SilentlyContinue)) {
    if (Get-Module -Name PatchManagementSupportTools -ListAvailable) {
        Import-Module PatchManagementSupportTools -ErrorAction SilentlyContinue
    }
}

if (-not (Get-Command -Name Get-PatchTuesday -ErrorAction SilentlyContinue)) {
    throw "Get-PatchTuesday was not found. Install/import PatchManagementSupportTools or adjust the script."
}

if (Get-Command -Name Get-CMModule -ErrorAction SilentlyContinue) {
    Get-CMModule | Out-Null
}

#########################################################
# Connect to CM PSDrive + main logic
#########################################################
$sitecode = Get-CMSiteCode -ComputerName $SiteServer
$setSiteCode = "$sitecode`:";
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
    # Patch Tuesday gating (robust date compare)
    #########################################################
    $monthname = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month)
    $year      = (Get-Date).Year

    $today = (Get-Date).Date  # Date-only

    $patchTuesdayThisMonth = Get-PatchTuesday -Month $today.Month -Year $today.Year
    $runDate = $patchTuesdayThisMonth.Date.AddDays($DaysAfterPatchTuesdayToReport)

    if ($DisableReportMonths -and ($today.Month -in $DisableReportMonths)) {
        Write-Log -LogString "$scriptname - This month ($($today.Month)) is skipped (DisableReportMonths)."
        Write-Log -LogString "$scriptname - Script exit!"
        return
    }

    if ($today -ne $runDate) {
        Write-Log -LogString ("========= {0} - Date not equal PatchTuesday {1} and it's now {2}. This report will run {3} ========" -f `
            $scriptname, $patchTuesdayThisMonth.ToShortDateString(), $today.ToShortDateString(), $runDate.ToString('yyyy-MM-dd'))
        Write-Log -LogString "========================== $scriptname - Script exit! =========================="
        return
    }

    #########################################################
    # Gather updates from Software Update Group
    #########################################################
    Write-Log -LogString "=====================Processing Software Update Group $UpdateGroupName=========================="

    $result = @()
    $updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName

    foreach ($item in $updates) {
        $result += New-Object psobject -Property @{ 
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
    # Build HTML report (PSWriteHTML) - aligned with Send-CheckDPStatus.ps1
    #########################################################
    # Keep the original "pre" and "post" text semantics/wording
    $preText  = "<h1>Windows updates $monthname $year</h1><br><p>The following updates from Microsoft are deployed this month.</p>"
    $postText = "<p>Raport created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>"

    # Normalize to array for table rendering
    $UpdatesFound = @($UpdatesFound)

    $reportTitle = "WindowsUpdate $MailCustomer $monthname $year"

    $html = New-HTML -TitleText $reportTitle -Online {
        New-HTMLTag -Tag 'style' {
@"
body { font-family: Arial, Helvetica, sans-serif; }
.small { font-size: 11px; color: #333; }
"@
        }

        # Same embedded image as Send-CheckDPStatus.ps1
        new-HTMLtag -Tag 'img alt="My image" src="data:image/png;base64,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\n"';

        # Pre content
        New-HTMLSection -HeaderText "Windows updates" -HeaderTextSize 20 -HeaderBackGroundColor Darkblue {
            New-HTMLTag -Tag 'div' {
                $preText
            }
        }

        # Updates table
        New-HTMLSection -HeaderText "Updates ($($UpdatesFound.Count))" -HeaderTextSize 20 -HeaderBackGroundColor Darkblue {
            if ($UpdatesFound.Count -eq 0) {
                New-HTMLText -Text 'No updates found within the configured LimitDays window.' -Color Gray
            }
            else {
                $tableData = $UpdatesFound | Select-Object ArticleID, Title, LocalizedDescription, DatePosted, Deployed, URL, Severity
                New-HTMLTable -DataTable $tableData
            }
        }

        # Post content
        New-HTMLSection -HeaderText "Report info" -HeaderTextSize 20 -HeaderBackGroundColor Darkblue {
            New-HTMLTag -Tag 'div' {
                $postText
            }
        }

        New-HTMLTag -Tag 'p' -Attributes @{ class = 'small' } {
            "Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b>"
        }
    }

    #########################################################
    # Send mail (Send-MailKitMessage)
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
    $HTMLBody = [string]$html

    $AttachmentList = New-Object 'System.Collections.Generic.List[string]'

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

    Write-Log -LogString ("========================== {0} - Mail on it\n's way to {1}" -f $scriptname, ($MailTo -join ', '))
    Write-Log -LogString "========================== $scriptname - Script exit! =========================="
}
finally {
    Pop-Location
}