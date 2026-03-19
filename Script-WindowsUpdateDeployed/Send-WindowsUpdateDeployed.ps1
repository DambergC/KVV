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
    # Build HTML report (PSWriteHTML; replaces ConvertTo-Html)
    #########################################################
    # Keep the original "pre" and "post" text semantics/wording
    $preText  = "<h1>Windows updates $monthname $year</h1><br><p>The following updates from Microsoft are deployed this month.</p>"
    $postText = "<p>Raport created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>"

    # Normalize to array for table rendering
    $UpdatesFound = @($UpdatesFound)

    $UpdatesFoundHtml = New-HTML -Title "Downloaded patches" -Online {
        # Inline CSS kept close to the previous ConvertTo-Html -Head style
        New-HTMLTag -Tag 'style' {
@"
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
body { background-color: lightgray; }
table {
    border: 1px solid black;
    border-collapse: collapse;
}
td {
    border: 1px solid black;
    padding: 5px;
    background-color: #E0F3F7;
}
"@
        }

        # Pre content
        New-HTMLSection {
            New-HTMLTag -Tag 'div' {
                $preText
            }
        }

        # Table content
        New-HTMLSection -BackgroundColor AirForceBlue {
            if ($UpdatesFound.Count -eq 0) {
                New-HTMLText -Text "No updates found within the configured LimitDays window." -Color Red -FontSize 14
            }
            else {
                # Keep same columns / order as the objects we create above
                $tableData = $UpdatesFound | Select-Object ArticleID, Title, LocalizedDescription, DatePosted, Deployed, URL, Severity
                New-HTMLTable -DataTable $tableData
            }
        }

        # Post content
        New-HTMLSection {
            New-HTMLTag -Tag 'div' {
                $postText
            }
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
    $HTMLBody = [string]$UpdatesFoundHtml

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

    Write-Log -LogString ("========================== {0} - Mail on it�s way to {1}" -f $scriptname, ($MailTo -join ', '))
    Write-Log -LogString "========================== $scriptname - Script exit! =========================="
}
finally {
    Pop-Location
}
