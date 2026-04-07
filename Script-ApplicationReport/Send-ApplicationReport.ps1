<#
.SYNOPSIS
  Sends per-application SCCM software installation summaries as HTML emails.

.DESCRIPTION
  - Reads config from Send-ApplicationReport.xml
  - Uses dbatools (Invoke-DbaQuery) to query the CM_ database
  - Builds HTML with PSWriteHTML
  - Sends HTML body via Send-MailKitMessage (Send-MailKitMessage module)
  - Sends one email per recipient

  Per-application schedule:
    <Application Name="Google Chrome" SendDays="Fri"> ... </Application>
    Supported days: Mon,Tue,Wed,Thu,Fri,Sat,Sun

  Deployment target reporting (collection-based, supports multiple collections):
    <Application ... TargetType="Device|User">
      <TargetCollections>
        <CollectionId>KV10039C</CollectionId>
        ...
      </TargetCollections>
    </Application>

  NEW (collection names):
  - The targets table will include CollectionName for each CollectionId (from v_Collection).
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$MailOnly,
    [switch]$AttachHTML
)

$scriptversion = '2.6'
$scriptname    = $MyInvocation.MyCommand.Name

Write-Host "Script: $scriptname  Version: $scriptversion"
if ($DryRun)     { Write-Host "Mode: DRY RUN (no emails will be sent, HTML will be opened)" }
if ($MailOnly)   { Write-Host "Mode: MAIL ONLY (no HTML files will be written)" }
if ($AttachHTML) { Write-Host "Mode: ATTACH HTML ENABLED (per-recipient attach flag from XML will be honored)" }

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$LogString,
        [Parameter()][ValidateSet('INFO','WARNING','ERROR','SUCCESS')][string]$Severity = 'INFO'
    )
    Write-Host "[$Severity] $LogString"
}

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
function Import-RequiredModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    try {
        if (-not (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)) {
            if (-not (Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue)) {
                Write-Log -LogString "Required module '$ModuleName' not found. Please install it first." -Severity "ERROR"
                return $false
            }
            Write-Log -LogString "Importing module '$ModuleName'" -Severity "INFO"
            Import-Module $ModuleName -ErrorAction Stop
            Write-Log -LogString "Successfully imported module '$ModuleName'" -Severity "INFO"
        }
        else {
            Write-Log -LogString "Module '$ModuleName' already loaded" -Severity "INFO"
        }
        return $true
    }
    catch {
        Write-Log -LogString "Failed to import module '$ModuleName'. Error: $_" -Severity "ERROR"
        return $false
    }
}

$requiredModules  = @("send-mailkitmessage", "PSWriteHTML", "dbatools")
$allModulesLoaded = $true

foreach ($module in $requiredModules) {
    $moduleLoaded     = Import-RequiredModule -ModuleName $module
    $allModulesLoaded = $allModulesLoaded -and $moduleLoaded
}

if (-not $allModulesLoaded) {
    Write-Log -LogString "Not all required modules could be loaded. Exiting script." -Severity "ERROR"
    exit 1
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

$querySummary = @"
SELECT
    CASE
        WHEN CS.Model0 LIKE '%Laptop%'         THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Notebook%'       THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Book%'           THEN 'Laptop'
        WHEN CS.Model0 LIKE '%EliteBook%'      THEN 'Laptop'
        WHEN CS.Model0 LIKE '%ProBook%'        THEN 'Laptop'
        WHEN CS.Model0 LIKE '%ThinkPad%'       THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Latitude%'       THEN 'Laptop'
        WHEN CS.Model0 LIKE '%XPS 13%' OR CS.Model0 LIKE '%XPS 15%' OR CS.Model0 LIKE '%XPS 17%' THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Surface Laptop%' THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Workstation%'    THEN 'Desktop'
        WHEN CS.Model0 LIKE '%Desktop%'        THEN 'Desktop'
        WHEN CS.Model0 LIKE '%OptiPlex%'       THEN 'Desktop'
        WHEN CS.Model0 LIKE '%ProDesk%'        THEN 'Desktop'
        WHEN CS.Model0 LIKE '%EliteDesk%'      THEN 'Desktop'
        WHEN CS.Model0 LIKE '%ThinkCentre%'    THEN 'Desktop'
        WHEN CS.Model0 LIKE '%Surface Studio%' THEN 'Desktop'
        ELSE 'Unknown'
    END AS [FormFactor],
    COUNT(DISTINCT SYS.ResourceID) AS [Device_Count]
FROM
    v_R_System AS SYS
    JOIN v_GS_COMPUTER_SYSTEM AS CS ON SYS.ResourceID = CS.ResourceID
    JOIN v_GS_OPERATING_SYSTEM AS OS ON SYS.ResourceID = OS.ResourceID
WHERE
    SYS.Client0 = 1
    AND OS.Caption0 LIKE '%Windows 11%'
    AND CS.Manufacturer0 NOT LIKE '%VMware%'
    AND CS.Manufacturer0 NOT LIKE '%VirtualBox%'
    AND NOT (CS.Manufacturer0 LIKE '%Microsoft%' AND CS.Model0 LIKE '%Virtual%')
GROUP BY
    CASE
        WHEN CS.Model0 LIKE '%Laptop%'         THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Notebook%'       THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Book%'           THEN 'Laptop'
        WHEN CS.Model0 LIKE '%EliteBook%'      THEN 'Laptop'
        WHEN CS.Model0 LIKE '%ProBook%'        THEN 'Laptop'
        WHEN CS.Model0 LIKE '%ThinkPad%'       THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Latitude%'       THEN 'Laptop'
        WHEN CS.Model0 LIKE '%XPS 13%' OR CS.Model0 LIKE '%XPS 15%' OR CS.Model0 LIKE '%XPS 17%' THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Surface Laptop%' THEN 'Laptop'
        WHEN CS.Model0 LIKE '%Workstation%'    THEN 'Desktop'
        WHEN CS.Model0 LIKE '%Desktop%'        THEN 'Desktop'
        WHEN CS.Model0 LIKE '%OptiPlex%'       THEN 'Desktop'
        WHEN CS.Model0 LIKE '%ProDesk%'        THEN 'Desktop'
        WHEN CS.Model0 LIKE '%EliteDesk%'      THEN 'Desktop'
        WHEN CS.Model0 LIKE '%ThinkCentre%'    THEN 'Desktop'
        WHEN CS.Model0 LIKE '%Surface Studio%' THEN 'Desktop'
        ELSE 'Unknown'
    END
ORDER BY
    [FormFactor];
"@

# ------------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------------
function Get-NormalizedSendDays {
    param([Parameter(Mandatory)][string]$SendDays)

    $valid = @('Mon','Tue','Wed','Thu','Fri','Sat','Sun')

    $days = $SendDays.Split(',') |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' } |
        ForEach-Object {
            $d = $_
            if ($d.Length -ge 3) { ($d.Substring(0,1).ToUpper() + $d.Substring(1,2).ToLower()) } else { $d }
        }

    $invalid = $days | Where-Object { $valid -notcontains $_ }

    [pscustomobject]@{
        ValidDays   = $valid
        AllowedDays = @($days)
        InvalidDays = @($invalid)
    }
}

# Robust XML parsing using XPath (less fragile than $node.TargetCollections.CollectionId)
function Get-XmlCollectionIds {
    param([Parameter(Mandatory)][System.Xml.XmlElement]$AppNode)

    $ids = @()

    # Preferred: <TargetCollections><CollectionId>...</CollectionId></TargetCollections>
    $nodes = $AppNode.SelectNodes('./TargetCollections/CollectionId')
    foreach ($n in @($nodes)) {
        $id = $n.InnerText
        if ($id) {
            $id = $id.Trim()
            if (-not [string]::IsNullOrWhiteSpace($id)) { $ids += $id }
        }
    }

    # Optional back-compat: <TargetCollectionId>KV123</TargetCollectionId>
    $single = $AppNode.SelectSingleNode('./TargetCollectionId')
    if ($single -and $single.InnerText) {
        $id = $single.InnerText.Trim()
        if (-not [string]::IsNullOrWhiteSpace($id)) { $ids += $id }
    }

    @($ids | Select-Object -Unique)
}

function ConvertTo-SqlInList {
    param([Parameter(Mandatory)][string[]]$Values)

    $quoted = $Values | ForEach-Object { "'" + ($_.Replace("'", "''")) + "'" }
    ($quoted -join ',')
}

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

    # -------------------------------------------------------------------------
    # Per-application schedule gate (SendDays="Mon,Tue,Fri")
    # -------------------------------------------------------------------------
    $today = (Get-Date).ToString('ddd', [System.Globalization.CultureInfo]::InvariantCulture) # Mon/Tue/...

    $sendDaysAttr = $app.SendDays
    if (-not [string]::IsNullOrWhiteSpace($sendDaysAttr)) {
        $sched = Get-NormalizedSendDays -SendDays $sendDaysAttr

        if ($sched.InvalidDays.Count -gt 0) {
            Write-Warning "Application '$appName' has invalid SendDays value(s): $($sched.InvalidDays -join ', '). Valid: $($sched.ValidDays -join ', ')"
            continue
        }

        if ($sched.AllowedDays.Count -gt 0 -and ($sched.AllowedDays -notcontains $today)) {
            Write-Host "Skipping '$appName' today ($today). Allowed days: $($sched.AllowedDays -join ', ')"
            continue
        }
    }

    Write-Host "=== Processing application '$appName' (filter: $productFilter) ==="

    # -------------------------------------------------------------------------
    # Deployment targets (unique) from one or more collections (Device/User)
    # + include collection name in the output table
    # -------------------------------------------------------------------------
    $targetType = 'Device'
    if ($app.TargetType -and -not [string]::IsNullOrWhiteSpace($app.TargetType.ToString())) {
        $targetType = $app.TargetType.ToString().Trim()
    }

    $collectionIds = Get-XmlCollectionIds -AppNode $app

    $targetsByCollection = @()
    $uniqueTargets = $null

    if ($collectionIds.Count -gt 0) {
        $isUser = $false
        if ($targetType.ToLower() -eq 'user') { $isUser = $true }

        $membershipView = if ($isUser) { 'v_FullCollectionMembership_User' } else { 'v_FullCollectionMembership' }
        $inList = ConvertTo-SqlInList -Values $collectionIds

        # Join against v_Collection to get the collection name
        $qByCollection = @"
SELECT
    m.CollectionID AS CollectionId,
    c.Name         AS CollectionName,
    COUNT(1)       AS TargetCount
FROM $membershipView AS m
LEFT JOIN v_Collection AS c
    ON c.CollectionID = m.CollectionID
WHERE
    m.CollectionID IN ($inList)
GROUP BY
    m.CollectionID,
    c.Name
ORDER BY
    m.CollectionID;
"@

        $qUnique = @"
SELECT
    COUNT(DISTINCT m.ResourceID) AS UniqueTargets
FROM $membershipView AS m
WHERE
    m.CollectionID IN ($inList);
"@

        try {
            $targetsByCollection = @(Invoke-DbaQuery -SqlInstance $SqlServer -Database $Database -Query $qByCollection -ErrorAction Stop |
                Select-Object CollectionId, CollectionName, TargetCount)

            $u = Invoke-DbaQuery -SqlInstance $SqlServer -Database $Database -Query $qUnique -ErrorAction Stop
            $uniqueTargets = [int]$u.UniqueTargets
        }
        catch {
            Write-Warning "Failed to query deployment targets for '$appName' ($targetType). Collections: $($collectionIds -join ', '). Error: $($_.Exception.Message)"
            $targetsByCollection = @()
            $uniqueTargets = $null
        }
    }

    # -------------------------------------------------------------------------
    # App inventory query
    # -------------------------------------------------------------------------
    $query = $queryTemplate.Replace('[PRODUCT_FILTER]', $productFilter)

    try {
        $dt      = Invoke-DbaQuery -SqlInstance $SqlServer -Database $Database -Query $query -ErrorAction Stop
        $summary = Invoke-DbaQuery -SqlInstance $SqlServer -Database $Database -Query $querySummary -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to run query for application '$appName': $($_.Exception.Message)"
        continue
    }

    $dt = @($dt)

    if ($dt.Count -gt 0) {
        $dt      = $dt | Select-Object Publisher, ProductName, ProductVersion, InstallCount
        $summary = $summary | Select-Object FormFactor, Device_count
    }

    # -------------------------------------------------------------------------
    # Build HTML
    # -------------------------------------------------------------------------
    $reportTitle = "$appName Software Installation Summary"
    $now         = Get-Date -Format "yyyy-MM-dd HH:mm"

    $html = New-HTML -TitleText $reportTitle -Online {

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
    max-width: 900px;
}
.header-block div {
    margin: 2px 0;
}
.header-label {
    font-weight: bold;
    text-align: center;
}
"@
        }

        # Optional: re-add your base64 logo line here if you want it embedded
        # new-HTMLtag -Tag 'img alt="My image" src="data:image/png;base64,...."'

        $targetLabel = if ($targetType.ToLower() -eq 'user') { 'Users' } else { 'Devices' }

        # Header section
        New-HTMLSection -HeaderTextAlignment center -HeaderTextSize 20 -HeaderBackGroundColor Darkblue -HeaderText $reportTitle {

            New-HTMLTag -Tag 'div' -Attributes @{ class = 'header-block' } {

                New-HTMLTag -Tag 'div' { "<span class='header-label'>Report generated:</span> $($now)" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Customer:</span> $MailCustomer" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Filter:</span> ProductName LIKE '$productFilter'" }

                if ($dt.Count -gt 0) {
                    $totalInstalls = ($dt | Measure-Object -Property InstallCount -Sum).Sum
                    $totaldevices  = ($summary | Measure-Object -Property device_count -Sum).Sum
                    New-HTMLTag -Tag 'div' {
                        "<span class='header-label'>Total installations (sum of InstallCount):</span> <b>$totalInstalls</b> of total <b>$totaldevices</b> Devices"
                    }
                }
                else {
                    New-HTMLTag -Tag 'div' { "<span class='header-label'>Total installations (sum of InstallCount):</span> 0" }
                }

                # Deployment target summary (UniqueTargets across collections)
                if ($collectionIds.Count -gt 0) {
                    if ($null -ne $uniqueTargets) {
                        New-HTMLTag -Tag 'div' {
                            "<span class='header-label'>Deployment unique targets ($targetLabel):</span> <b>$uniqueTargets</b>"
                        }
                    }
                    New-HTMLTag -Tag 'div' {
                        "<span class='header-label'>Target collections ($targetType):</span> $($collectionIds -join ', ')"
                    }
                }
                else {
                    New-HTMLTag -Tag 'div' {
                        "<span class='header-label'>Deployment targets:</span> No TargetCollections configured"
                    }
                }
            }
        }

        # Data section: software table + (ALWAYS UNDER) deployment targets table
        New-HTMLSection -HeaderBackGroundColor darkblue {

            if ($dt.Count -eq 0) {
                New-HTMLText -Text "No matching installations found." -Color Red -FontSize 14
            }
            else {
                New-HTMLTable -DataTable $dt
            }

            # Always under software table
            New-HTMLText -Text "Deployment targets per collection ($targetLabel)" -Color White -FontSize 14

            if ($collectionIds.Count -gt 0) {
                if ($targetsByCollection.Count -gt 0) {
                    New-HTMLTable -DataTable $targetsByCollection
                }
                else {
                    New-HTMLText -Text "No collection membership rows returned (check CollectionId / TargetType)." -Color Yellow -FontSize 12
                }
            }
            else {
                New-HTMLText -Text "Not configured for this application." -Color Yellow -FontSize 12
            }
        }
    }

    # -------------------------------------------------------------------------
    # Handle HTML file output (skipped when -MailOnly is used)
    # -------------------------------------------------------------------------
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
            $tempFile = [System.IO.Path]::GetTempFileName().Replace('.tmp', '.html')
            $html | Out-File -FilePath $tempFile -Encoding UTF8
            Start-Process $tempFile
        }
        continue
    }

    # -------------------------------------------------------------------------
    # Recipients
    # -------------------------------------------------------------------------
    $recipientNodes = $app.Recipients.Recipient
    if (-not $recipientNodes) {
        Write-Warning "Application '$appName' has no <Recipients><Recipient> entries. Skipping email."
        continue
    }

    $recipients = @()
    foreach ($r in $recipientNodes) {
        if (-not $r.email) { continue }

        $attachAttr = $r.attach
        $attachFlag = $false
        if ($attachAttr -and $attachAttr.ToString().ToLower() -eq 'true') {
            $attachFlag = $true
        }

        $recipients += [pscustomobject]@{
            Email         = $r.email
            AttachFromXml = $attachFlag
        }
    }

    if ($recipients.Count -eq 0) {
        Write-Warning "Application '$appName' has Recipient nodes but no 'email' values. Skipping email."
        continue
    }

    Write-Host "Recipients for '$appName': $($recipients.Email -join ', ')"

    # -------------------------------------------------------------------------
    # Mail prepare (include UniqueTargets in subject when available)
    # -------------------------------------------------------------------------
    $fromAddress = New-Object MimeKit.MailboxAddress ('', $MailFrom)

    $targetLabelShort = if ($targetType.ToLower() -eq 'user') { 'Users' } else { 'Devices' }
    if ($null -ne $uniqueTargets) {
        $Subject = "$appName Software Installation Summary - $now (Targets: $uniqueTargets $targetLabelShort)"
    } else {
        $Subject = "$appName Software Installation Summary - $now"
    }

    $baseAttachment = $null
    if ($htmlFile -and (Test-Path -LiteralPath $htmlFile)) {
        $baseAttachment = $htmlFile
    }

    # -------------------------------------------------------------------------
    # Send one email per recipient
    # -------------------------------------------------------------------------
    foreach ($rec in $recipients) {
        $addr          = $rec.Email
        $attachFromXml = $rec.AttachFromXml

        if ([string]::IsNullOrWhiteSpace($addr)) { continue }

        Write-Host "Sending mail for '$appName' to $addr (attach from XML: $attachFromXml) ..."

        $recipientAddress = New-Object MimeKit.MailboxAddress ('', $addr)

        $params = @{
            SMTPServer    = $SMTP
            Port          = $MailPortnumber
            From          = $fromAddress
            RecipientList = $recipientAddress
            Subject       = $Subject
            HTMLBody      = $html
        }

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
