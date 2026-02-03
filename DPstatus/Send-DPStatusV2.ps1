<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Checks MECM Distribution Points reachability and toggles Maintenance Mode / DP Group membership accordingly,
   then emails an HTML report.

.DESCRIPTION
   Intended to run as a scheduled task on the site server.

   Flow:
   - Import required modules (ConfigurationManager, Send-MailKitMessage, PSWriteHTML (optional))
   - Connect to the MECM PSDrive
   - Enumerate DPs
   - Check reachability (ICMP ping by default)
   - If DP is offline and not already in maintenance -> Enable maintenance + move to Maintenance group
   - If DP is online and in maintenance -> Disable maintenance + move back to Production group
   - Produce HTML report and send email
   - Write log file entries for all actions

.NOTES
   - Scripts are offered AS IS with no warranty.
   - Test in non-production first.

-------------------------------------------------------------------------------------------------------------------------
#>

[CmdletBinding()]
param(
    # Site server used for retrieving the CM SiteCode via WMI/CIM
    [Parameter()]
    [string]$SiteServer = 'vntsql0299',

    # DP groups
    [Parameter()]
    [string]$DPMaintGroup = 'Maintenance',

    [Parameter()]
    [string]$DPProdGroup = 'VS DPs',

    # Reachability test
    [Parameter()]
    [ValidateSet('ICMP')]
    [string]$ReachabilityMethod = 'ICMP',

    [Parameter()]
    [int]$PingCount = 1,

    # Mail settings
    [Parameter()]
    [string]$MailFrom = 'no-reply@kvv.se',

    [Parameter()]
    [string[]]$MailTo = @(
        'christian.damberg@kriminalvarden.se',
        'Joakim.Stenqvist@kriminalvarden.se',
        'magnus.jonsson6@kriminalvarden.se',
        'Keiarash.Naderifarsani@kriminalvarden.se',
        'it.andralinjen@kriminalvarden.se'
    ),

    [Parameter()]
    [string]$MailSMTP = 'smtp.kvv.se',

    [Parameter()]
    [int]$MailPortnumber = 25,

    [Parameter()]
    [switch]$UseSecureConnectionIfAvailable,

    # Log file
    [Parameter()]
    [string]$Logfile = "G:\Scripts\Logfiles\DPLogfile.log",

    # Report cosmetics
    [Parameter()]
    [string]$CustomerName = 'Kriminalvården'
)

$ErrorActionPreference = 'Stop'

$scriptname = $MyInvocation.MyCommand.Name
$today = (Get-Date).ToString("yyyy-MM-dd")

#########################################################
# Logging
#########################################################
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$LogString
    )

    try {
        $stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
        $logMessage = "$stamp $LogString"
        Add-Content -Path $Logfile -Value $logMessage
    } catch {
        # If logging fails, don't kill the whole run; write to host as fallback
        Write-Warning "Failed to write to log file '$Logfile': $($_.Exception.Message)"
    }
}

#########################################################
# SiteCode helper
#########################################################
function Get-CMSiteCode {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    # Prefer CIM over deprecated Get-WmiObject
    $prov = Get-CimInstance -Namespace "root\SMS" -ClassName SMS_ProviderLocation -ComputerName $ComputerName |
        Select-Object -First 1 -ExpandProperty SiteCode

    if (-not $prov) {
        throw "Unable to determine SiteCode from SMS_ProviderLocation on '$ComputerName'."
    }
    return $prov
}

#########################################################
# Module imports
#########################################################
function Ensure-Module {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [ScriptBlock]$ImportScript
    )

    if (-not (Get-Module -Name $Name -ListAvailable)) {
        throw "Required module '$Name' is not available on this machine."
    }

    if (-not (Get-Module -Name $Name)) {
        if ($ImportScript) {
            & $ImportScript
        } else {
            Import-Module -Name $Name
        }
    }
}

Write-Log -LogString "$scriptname - Script start"

# ConfigurationManager is special: often imported via $Env:SMS_ADMIN_UI_PATH
if (-not (Get-Module -Name ConfigurationManager)) {
    if (-not $Env:SMS_ADMIN_UI_PATH) {
        throw "SMS_ADMIN_UI_PATH is not set. Run on a machine with the ConfigMgr console installed."
    }
    $cmModulePath = Join-Path -Path ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5)) -ChildPath 'ConfigurationManager.psd1'
    Import-Module $cmModulePath
}

Ensure-Module -Name 'Send-MailKitMessage' -ImportScript { Import-Module Send-MailKitMessage }

# PSWriteHTML is optional (we don't strictly need it), but keep compatibility
if (Get-Module -Name PSWriteHTML -ListAvailable) {
    if (-not (Get-Module -Name PSWriteHTML)) { Import-Module PSWriteHTML }
}

# Some environments define Get-CMModule; some don't. Use if available.
if (Get-Command -Name Get-CMModule -ErrorAction SilentlyContinue) {
    Get-CMModule | Out-Null
}

#########################################################
# Connect to CM PSDrive
#########################################################
$sitecode = Get-CMSiteCode -ComputerName $SiteServer
$setSiteCode = "$sitecode`:"  # e.g. "ABC:"
Write-Log -LogString "$scriptname - SiteCode resolved: $sitecode"

Push-Location
try {
    Set-Location $setSiteCode
} catch {
    Pop-Location
    throw "Failed to Set-Location to ConfigMgr PSDrive '$setSiteCode'. Error: $($_.Exception.Message)"
}

#########################################################
# Gather DPs
#########################################################
Write-Host "Get data about DPs" -ForegroundColor Green
$dpList = Get-CMDistributionPoint -SiteCode $sitecode |
    Select-Object -ExpandProperty NetworkOSPath |
    Sort-Object

if (-not $dpList) {
    Write-Log -LogString "$scriptname - No Distribution Points found."
    throw "No Distribution Points found in site '$sitecode'."
}

# Normalize DP names (strip leading \\ and upper-case)
$dpNames = foreach ($path in $dpList) {
    (($path -replace '\\\\', '') -replace '\\', '').ToUpperInvariant()
}

#########################################################
# Reachability check
#########################################################
$succeededDPs = New-Object System.Collections.Generic.List[string]
$failedDPs    = New-Object System.Collections.Generic.List[string]

foreach ($dp in $dpNames) {
    $online = $false

    switch ($ReachabilityMethod) {
        'ICMP' {
            $online = Test-Connection -ComputerName $dp -Count $PingCount -Quiet -ErrorAction SilentlyContinue
        }
    }

    if ($online) {
        Write-Host "$dp is online" -ForegroundColor Green
        $succeededDPs.Add($dp) | Out-Null
    } else {
        Write-Host "$dp is not online" -ForegroundColor Red
        Write-Log -LogString "$dp not reachable via $ReachabilityMethod"
        $failedDPs.Add($dp) | Out-Null
    }
}

$SucceededDPscount = $succeededDPs.Count
$FailedDpscount = $failedDPs.Count
$TotalDPscount = $SucceededDPscount + $FailedDpscount

#########################################################
# Helpers for DP actions
#########################################################
function New-ResultRow {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][string]$Description,
        [Parameter(Mandatory)][string]$Status,
        [Parameter()][string]$Action,
        [Parameter()][string]$Notes
    )

    [pscustomobject]@{
        Timestamp   = (Get-Date)
        Name        = $Name
        Description = $Description
        Status      = $Status
        Action      = $Action
        Notes       = $Notes
    }
}

$resultOfflineActions = New-Object System.Collections.Generic.List[object]
$resultRestoreActions = New-Object System.Collections.Generic.List[object]
$resultNoChange       = New-Object System.Collections.Generic.List[object]
$resultErrors         = New-Object System.Collections.Generic.List[object]

#########################################################
# Process OFFLINE DPs
#########################################################
if ($failedDPs.Count -gt 0) {
    $i = 0
    foreach ($failedDP in $failedDPs) {
        $i++
        Write-Host "Processing offline... $i of $($failedDPs.Count): $failedDP" -ForegroundColor Yellow

        try {
            $dpInfo = Get-CMDistributionPointInfo -SiteSystemServerName $failedDP
            $desc = $dpInfo.Description

            # MaintenanceMode seems to be 0/1 in many environments
            $isInMaintenance = [bool]$dpInfo.MaintenanceMode

            if (-not $isInMaintenance) {
                Write-Log -LogString "$failedDP offline -> enabling Maintenance Mode and moving to group '$DPMaintGroup' (removing from '$DPProdGroup')"

                Set-CMDistributionPoint -SiteSystemServerName $failedDP -EnableMaintenanceMode $true -Force
                Add-CMDistributionPointToGroup -DistributionPointName $failedDP -DistributionPointGroupName $DPMaintGroup
                Remove-CMDistributionPointFromGroup -DistributionPointName $failedDP -DistributionPointGroupName $DPProdGroup -Force

                $resultOfflineActions.Add((New-ResultRow -Name $failedDP -Description $desc -Status 'Offline' -Action 'Enabled maintenance; moved to Maintenance group' -Notes 'DP not reachable')) | Out-Null
            } else {
                # Already in maintenance; no change
                $resultNoChange.Add((New-ResultRow -Name $failedDP -Description $desc -Status 'Offline' -Action 'No change' -Notes 'Already in maintenance')) | Out-Null
            }
        } catch {
            $msg = $_.Exception.Message
            Write-Host "ERROR: Failed processing $failedDP. $msg" -ForegroundColor Red
            Write-Log -LogString "ERROR: Failed processing offline DP '$failedDP'. $msg"

            $resultErrors.Add((New-ResultRow -Name $failedDP -Description '' -Status 'Offline' -Action 'Error' -Notes $msg)) | Out-Null
        }
    }
}

#########################################################
# Process ONLINE DPs
#########################################################
if ($succeededDPs.Count -gt 0) {
    $i = 0
    foreach ($okDP in $succeededDPs) {
        $i++
        Write-Host "Processing online... $i of $($succeededDPs.Count): $okDP" -ForegroundColor Yellow

        try {
            $dpInfo = Get-CMDistributionPointInfo -SiteSystemServerName $okDP
            $desc = $dpInfo.Description
            $isInMaintenance = [bool]$dpInfo.MaintenanceMode

            if ($isInMaintenance) {
                Write-Host "$okDP is online but in Maintenance Mode -> restoring and moving to '$DPProdGroup'..." -ForegroundColor DarkCyan
                Write-Log -LogString "$okDP online + maintenance -> disabling Maintenance Mode and moving to group '$DPProdGroup' (removing from '$DPMaintGroup')"

                Set-CMDistributionPoint -SiteSystemServerName $okDP -EnableMaintenanceMode $false -Force
                Add-CMDistributionPointToGroup -DistributionPointName $okDP -DistributionPointGroupName $DPProdGroup
                Remove-CMDistributionPointFromGroup -DistributionPointName $okDP -DistributionPointGroupName $DPMaintGroup -Force

                $resultRestoreActions.Add((New-ResultRow -Name $okDP -Description $desc -Status 'Online' -Action 'Disabled maintenance; moved to Production group' -Notes 'DP reachable again')) | Out-Null
            } else {
                $resultNoChange.Add((New-ResultRow -Name $okDP -Description $desc -Status 'Online' -Action 'No change' -Notes 'Online and not in maintenance')) | Out-Null
            }
        } catch {
            $msg = $_.Exception.Message
            Write-Host "ERROR: Failed processing $okDP. $msg" -ForegroundColor Red
            Write-Log -LogString "ERROR: Failed processing online DP '$okDP'. $msg"

            $resultErrors.Add((New-ResultRow -Name $okDP -Description '' -Status 'Online' -Action 'Error' -Notes $msg)) | Out-Null
        }
    }
}

########################################################
# Build HTML report
#########################################################
$header = @"
<style>
    body { background-color: lightgray; font-family: Arial, Helvetica, sans-serif; }
    h1,h2 { font-family: Arial, Helvetica, sans-serif; }
    table { border: 1px solid black; border-collapse: collapse; width: 100%; }
    th {
        color: White;
        font-size: 12px;
        border: 1px solid black;
        padding: 6px;
        background-color: Black;
        text-align: left;
    }
    td {
        border: 1px solid black;
        padding: 6px;
        background-color: #E0F3F7;
        font-size: 11px;
        vertical-align: top;
    }
    .summary p { font-size: 12px; }
    .muted { color: #333; font-size: 11px; }
</style>
"@

$pre = @"
<div class="summary">
  <h1>$CustomerName - Dagens körning - Kontroll av Distributionspunkter</h1>
  <p>Datum: <b>$today</b></p>
  <p>Antal Distributionspunkter: <b>$TotalDPscount</b></p>
  <p>Antal Distributionspunkter som svarar på ping: <b>$SucceededDPscount</b></p>
  <p>Antal Distributionspunkter som inte svarar på ping: <b>$FailedDpscount</b></p>
</div>
"@

$post = @"
<p class="muted">Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
"@

function Convert-SectionToHtml {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][System.Collections.IEnumerable]$Rows
    )

    $rowsList = @($Rows)
    if ($rowsList.Count -eq 0) {
        return "<h2>$Title</h2><p>Inga poster.</p>"
    }

    ($rowsList |
        Select-Object Timestamp, Name, Description, Status, Action, Notes |
        ConvertTo-Html -Fragment -PreContent "<h2>$Title</h2>")
}

$sections = @()
$sections += Convert-SectionToHtml -Title "Offline DPs - Åtgärdade" -Rows $resultOfflineActions
$sections += Convert-SectionToHtml -Title "Online DPs - Återställda" -Rows $resultRestoreActions
$sections += Convert-SectionToHtml -Title "Ingen ändring" -Rows $resultNoChange
$sections += Convert-SectionToHtml -Title "Fel" -Rows $resultErrors

$bodyHtml = @()
$bodyHtml += "<html><head>$header</head><body>"
$bodyHtml += $pre
$bodyHtml += ($sections -join "`n")
$bodyHtml += $post
$bodyHtml += "</body></html>"

$HTMLBody = [string]($bodyHtml -join "`n")

#########################################################
# Send email (Send-MailKitMessage)
#########################################################
$RecipientList = [MimeKit.InternetAddressList]::new()
foreach ($addr in $MailTo) {
    if ([string]::IsNullOrWhiteSpace($addr)) { continue }
    $RecipientList.Add([MimeKit.InternetAddress]$addr)
}

$Subject = [string]"Kontroll av Distributionspunkter - $today"

$Parameters = @{
    UseSecureConnectionIfAvailable = [bool]$UseSecureConnectionIfAvailable
    SMTPServer                    = $MailSMTP
    Port                          = $MailPortnumber
    From                          = [MimeKit.MailboxAddress]$MailFrom
    RecipientList                 = $RecipientList
    Subject                       = $Subject
    HTMLBody                      = $HTMLBody
    AttachmentList                = [System.Collections.Generic.List[string]]::new()
}

try {
    Send-MailKitMessage @Parameters
    Write-Log -LogString "$scriptname - Mail sent (or queued by SMTP) to: $($MailTo -join ', ')"
} catch {
    Write-Host "ERROR: Failed to send mail. $($_.Exception.Message)" -ForegroundColor Red
    Write-Log -LogString "ERROR: Failed to send mail. $($_.Exception.Message)"
}

#########################################################
# Cleanup
#########################################################
try {
    Set-Location $PSScriptRoot
} catch {
    # ignore
}

Pop-Location
Write-Log -LogString "$scriptname - Script end"
