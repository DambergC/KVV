<#
.SYNOPSIS
    Queries server patch compliance for one or more ConfigMgr collections
    against a Software Update Group and sends one email per collection.

    Each collection is evaluated on the day that matches its configured
    offset (in days) from the current month's Patch Tuesday.

.REQUIREMENTS
    Install-Module dbatools -Scope CurrentUser
    Install-Module Send-MailKitMessage -Scope CurrentUser

    Configuration file:
    D:\Scripts\ServerComplianceReport\ServerComplianceReport.config.xml

.NOTES
    Intended to be scheduled once per day.
    If no collection matches the configured Patch Tuesday offset today,
    the script exits without sending any email.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "D:\Scripts\ServerComplianceReport\ServerComplianceReport.config.xml",
    [switch]$KeepHtmlFile,
    [datetime]$AsOfDate = (Get-Date)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Logging

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logLine = "$timestamp [$Level] $Message"

    try {
        switch ($Level) {
            "WARNING" { Write-Warning $Message }
            "ERROR"   { Write-Error $Message -ErrorAction Continue }
            default   { Write-Host $logLine }
        }
    }
    catch {
        # Ignore console problems in scheduled mode.
    }

    try {
        $logDirectory = Split-Path -Path $LogPath -Parent

        if (
            -not [string]::IsNullOrWhiteSpace($logDirectory) -and
            -not (Test-Path -LiteralPath $logDirectory)
        ) {
            New-Item -Path $logDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        Add-Content -LiteralPath $LogPath -Value $logLine -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # Do not allow logging failures to stop the report.
    }
}

#endregion

function Get-ConfigValue {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlNode]$Xml,
        [Parameter(Mandatory)]
        [string]$XPath,
        [switch]$Required
    )

    $node = $Xml.SelectSingleNode($XPath)

    if ($null -eq $node) {
        if ($Required) {
            throw "Required configuration value is missing: $XPath"
        }
        return ""
    }

    $value = ([string]$node.InnerText).Trim()

    if ($Required -and [string]::IsNullOrWhiteSpace($value)) {
        throw "Required configuration value is empty: $XPath"
    }

    return $value
}

function Get-PatchTuesday {
    param(
        [Parameter(Mandatory)]
        [int]$Year,
        [Parameter(Mandatory)]
        [int]$Month
    )

    $firstOfMonth = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0

    $firstDayOfWeek = [int]$firstOfMonth.DayOfWeek
    $daysUntilFirstTuesday = (2 - $firstDayOfWeek + 7) % 7
    $firstTuesday = $firstOfMonth.AddDays($daysUntilFirstTuesday)
    $patchTuesday = $firstTuesday.AddDays(7)

    return $patchTuesday
}

function Get-CurrentPatchTuesday {
    param(
        [Parameter(Mandatory)]
        [datetime]$AsOfDate
    )

    $thisMonthPatchTuesday = Get-PatchTuesday -Year $AsOfDate.Year -Month $AsOfDate.Month

    if ($AsOfDate.Date -ge $thisMonthPatchTuesday.Date) {
        return $thisMonthPatchTuesday
    }

    $previousMonthDate = $AsOfDate.AddMonths(-1)
    return Get-PatchTuesday -Year $previousMonthDate.Year -Month $previousMonthDate.Month
}

#endregion

#region Startup and configuration

$LogPath = "D:\Scripts\ServerComplianceReport\LogFiles\ServerComplianceReport.log"

Write-Log "============================================================"
Write-Log "Server Compliance Report started."
Write-Log "Computer: $env:COMPUTERNAME"
Write-Log "User: $env:USERNAME"
Write-Log "Script path: $PSCommandPath"
Write-Log "Configuration path: $ConfigPath"
Write-Log "PowerShell version: $($PSVersionTable.PSVersion)"

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    Write-Log "Configuration file not found: $ConfigPath" "ERROR"
    throw "Configuration file not found: $ConfigPath"
}

try {
    [xml]$configXml = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop
    Write-Log "XML configuration loaded successfully."
}
catch {
    Write-Log "Could not load XML configuration: $($_.Exception.Message)" "ERROR"
    throw
}

$SqlInstance = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Database/SqlInstance" -Required
$Database    = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Database/DatabaseName" -Required

$SmtpServer   = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Email/SmtpServer" -Required
$SmtpPortText = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Email/SmtpPort" -Required
$From         = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Email/From" -Required
$SubjectBase  = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Email/Subject" -Required

$HtmlLogo     = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Email/HtmlLogo"
$LogoMimeType = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Email/HtmlLogoMimeType"

if ([string]::IsNullOrWhiteSpace($LogoMimeType)) {
    $LogoMimeType = "image/png"
}

$configuredLogPath = Get-ConfigValue -Xml $configXml -XPath "/Configuration/Logging/LogPath"

if (-not [string]::IsNullOrWhiteSpace($configuredLogPath)) {
    $LogPath = $configuredLogPath
}

$recipientNodes = $configXml.SelectNodes("/Configuration/Email/To/Recipient")

$To = @(
    foreach ($recipientNode in $recipientNodes) {
        $recipient = ([string]$recipientNode.InnerText).Trim()
        if (-not [string]::IsNullOrWhiteSpace($recipient)) {
            $recipient
        }
    }
)

if ($To.Count -eq 0) {
    throw "No email recipients are configured."
}

try {
    $SmtpPort = [int]$SmtpPortText
}
catch {
    Write-Log "SMTP port is not numeric." "ERROR"
    throw
}

if ($SmtpPort -lt 1 -or $SmtpPort -gt 65535) {
    throw "SMTP port must be between 1 and 65535."
}

$collectionNodes = $configXml.SelectNodes("/Configuration/Query/Collections/Collection")

if ($collectionNodes.Count -eq 0) {
    throw "No collections are configured under /Configuration/Query/Collections."
}

$patchTuesday = Get-CurrentPatchTuesday -AsOfDate $AsOfDate
$daysSincePatchTuesday = [int]([math]::Floor(($AsOfDate.Date - $patchTuesday.Date).TotalDays))

Write-Log "As-of date: $($AsOfDate.ToString('yyyy-MM-dd'))"
Write-Log "Current Patch Tuesday: $($patchTuesday.ToString('yyyy-MM-dd'))"
Write-Log "Days since Patch Tuesday: $daysSincePatchTuesday"

$AllCollections = @(
    foreach ($node in $collectionNodes) {
        $name = Get-ConfigValue -Xml $node -XPath "Name" -Required
        $title = Get-ConfigValue -Xml $node -XPath "Title" -Required
        $xDaysText = Get-ConfigValue -Xml $node -XPath "XDays" -Required
        $offsetText = Get-ConfigValue -Xml $node -XPath "DaysAfterPatchTuesday" -Required

        [pscustomobject]@{
            Name                  = $name
            Title                 = $title
            XDays                 = [int]$xDaysText
            DaysAfterPatchTuesday = [int]$offsetText
        }
    }
)

Write-Log "Configured collection count: $($AllCollections.Count)"

foreach ($c in $AllCollections) {
    Write-Log "Configured collection '$($c.Name)' -> SUG '$($c.Title)', XDays=$($c.XDays), DaysAfterPatchTuesday=$($c.DaysAfterPatchTuesday)"
}

$CollectionsToRun = @(
    $AllCollections | Where-Object {
        $_.DaysAfterPatchTuesday -eq $daysSincePatchTuesday
    }
)

Write-Log "Collections matching today's offset ($daysSincePatchTuesday): $($CollectionsToRun.Count)"

if ($CollectionsToRun.Count -eq 0) {
    Write-Log "No collections are scheduled to run today. Exiting without sending an email."
    Write-Output "No collections matched today's Patch Tuesday offset ($daysSincePatchTuesday day(s)). Nothing to report."
    return
}

foreach ($c in $CollectionsToRun) {
    Write-Log "Will run today: '$($c.Name)' (SUG '$($c.Title)')"
}

#endregion

#region Helper functions

function ConvertTo-HtmlEncodedText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-LogoHtml {
    param(
        [AllowNull()]
        [string]$LogoValue,
        [string]$DefaultMimeType = "image/png"
    )

    if ([string]::IsNullOrWhiteSpace($LogoValue)) {
        Write-Log "No logo configured." "WARNING"
        return ""
    }

    $LogoValue = $LogoValue.Trim()

    if ($LogoValue -match '^data:image/[^;]+;base64,') {
        $parts = $LogoValue -split ",", 2
        $mimeHeader = $parts[0].Trim()
        $base64Content = $parts[1] -replace "\s", ""

        if ([string]::IsNullOrWhiteSpace($base64Content)) {
            Write-Log "Logo Base64 content is empty." "WARNING"
            return ""
        }

        $dataUri = "$mimeHeader,$base64Content"

        return @"
<p style="margin:0 0 15px 0;">
    <img
        src="$dataUri"
        alt="Logo"
        width="200"
        style="display:block;width:200px;max-width:100%;height:auto;border:0;"
    />
</p>
"@
    }

    $base64Content = $LogoValue -replace "\s", ""

    if ($base64Content -match "^[A-Za-z0-9+/]*={0,2}$" -and ($base64Content.Length % 4 -eq 0)) {
        $dataUri = "data:$DefaultMimeType;base64,$base64Content"

        return @"
<p style="margin:0 0 15px 0;">
    <img
        src="$dataUri"
        alt="Logo"
        width="200"
        style="display:block;width:200px;max-width:100%;height:auto;border:0;"
    />
</p>
"@
    }

    try {
        $logoUri = [System.Uri]$LogoValue

        if ($logoUri.Scheme -in @("http", "https")) {
            $encodedLogoUrl = [System.Net.WebUtility]::HtmlEncode($logoUri.AbsoluteUri)

            return @"
<p style="margin:0 0 15px 0;">
    <img
        src="$encodedLogoUrl"
        alt="Logo"
        width="200"
        style="display:block;width:200px;max-width:100%;height:auto;border:0;"
    />
</p>
"@
        }
    }
    catch {
        # Not a valid external logo URL.
    }

    Write-Log "Configured logo is not valid Base64 or HTTP/HTTPS." "WARNING"
    return ""
}

function New-ComputerListHtml {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$ComputerNames,
        [Parameter(Mandatory)]
        [string]$EmptyMessage,
        [string]$CellStyle = "background-color:#ffffff;color:#000000;"
    )

    if ($ComputerNames.Count -eq 0) {
        return @"
<p style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#444444;">
    $EmptyMessage
</p>
"@
    }

    $rows = [System.Text.StringBuilder]::new()

    foreach ($name in $ComputerNames) {
        $nameHtml = ConvertTo-HtmlEncodedText -Value $name

        [void]$rows.AppendLine(@"
<tr>
    <td style="border:1px solid #d9d9d9;padding:6px 8px;$CellStyle">
        $nameHtml
    </td>
</tr>
"@)
    }

    return @"
<table
    role="presentation"
    cellspacing="0"
    cellpadding="0"
    border="0"
    style="border-collapse:collapse;width:100%;max-width:400px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222222;"
>
    <thead>
        <tr style="background-color:#1f4e78;color:#ffffff;">
            <th style="border:1px solid #d9d9d9;padding:6px 8px;text-align:left;">Computer name</th>
        </tr>
    </thead>
    <tbody>
        $($rows.ToString())
    </tbody>
</table>
"@
}

function Get-CollectionComplianceQuery {
    param(
        [Parameter(Mandatory)]
        [string]$CollectionName,
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [int]$XDays
    )

    return @"
DECLARE @CollectionName NVARCHAR(255) = N'$CollectionName';
DECLARE @Title NVARCHAR(255) = N'$Title';
DECLARE @XDays INT = $XDays;

IF OBJECT_ID('tempdb..#ClientCompliance') IS NOT NULL
    DROP TABLE #ClientCompliance;

WITH CollectionMembers AS
(
    SELECT ResourceID
    FROM v_ClientCollectionMembers
    WHERE CollectionID =
    (
        SELECT TOP 1 CollectionID
        FROM v_Collection
        WHERE Name = @CollectionName
    )
),
HealthyClients AS
(
    SELECT DISTINCT
        sys.ResourceID
    FROM v_R_System sys
    LEFT JOIN
    (
        SELECT
            ResourceID,
            MAX(AgentTime) AS LastAgentTime
        FROM v_AgentDiscoveries
        WHERE AgentName = 'Heartbeat Discovery'
        GROUP BY ResourceID
    ) hb
        ON hb.ResourceID = sys.ResourceID
    LEFT JOIN v_CH_ClientSummary cs
        ON cs.ResourceID = sys.ResourceID
    INNER JOIN CollectionMembers cm
        ON cm.ResourceID = sys.ResourceID
    WHERE
    (
        hb.LastAgentTime >= DATEADD(DAY,-@XDays,GETDATE())
        OR cs.LastOnline >= DATEADD(DAY,-@XDays,GETDATE())
        OR cs.LastDDR >= DATEADD(DAY,-@XDays,GETDATE())
    )
    AND ISNULL(sys.Obsolete0,0) <> 1
    AND ISNULL(sys.Decommissioned0,0) <> 1
    AND sys.Client0 = 1
    AND sys.Active0 = 1
    AND sys.Client_Type0 = 1
),
SUG AS
(
    SELECT TOP 1 CI_ID
    FROM v_AuthListInfo
    WHERE Title = @Title
),
SUGUpdates AS
(
    SELECT DISTINCT ui.CI_ID
    FROM v_UpdateInfo ui
    WHERE ui.CI_ID IN
    (
        SELECT BundledCI_ID
        FROM v_BundledConfigurationItems
        WHERE CI_ID = (SELECT CI_ID FROM SUG)
    )
),
ClientCompliance AS
(
    SELECT
        rs.Name0 AS ComputerName,
        rs.ResourceID,

        CASE
            WHEN MAX(CASE WHEN ucs.Status = 2 THEN 1 ELSE 0 END) = 1
                THEN 'Non-Compliant'

            WHEN MAX(CASE WHEN ucs.Status = 0 THEN 1 ELSE 0 END) = 1
                THEN 'Unknown'

            ELSE 'Compliant'
        END AS ComplianceState

    FROM HealthyClients hc

    INNER JOIN v_R_System rs
        ON rs.ResourceID = hc.ResourceID

    LEFT JOIN v_Update_ComplianceStatusAll ucs
        ON ucs.ResourceID = hc.ResourceID
        AND ucs.CI_ID IN
        (
            SELECT CI_ID
            FROM SUGUpdates
        )

    GROUP BY
        rs.Name0,
        rs.ResourceID
)

SELECT *
INTO #ClientCompliance
FROM ClientCompliance;

----------------------------------------------------
-- Resultatset 1 - Summering
----------------------------------------------------
SELECT
    COUNT(*) AS Total,
    SUM(CASE WHEN ComplianceState = 'Compliant' THEN 1 ELSE 0 END) AS Compliant,
    SUM(CASE WHEN ComplianceState = 'Non-Compliant' THEN 1 ELSE 0 END) AS NonCompliant,
    SUM(CASE WHEN ComplianceState = 'Unknown' THEN 1 ELSE 0 END) AS UnknownClients
FROM #ClientCompliance;

----------------------------------------------------
-- Resultatset 2 - Non-Compliant datorer
----------------------------------------------------
SELECT
    ComputerName
FROM #ClientCompliance
WHERE ComplianceState = 'Non-Compliant'
ORDER BY ComputerName;

----------------------------------------------------
-- Resultatset 3 - Unknown datorer
----------------------------------------------------
SELECT
    ComputerName
FROM #ClientCompliance
WHERE ComplianceState = 'Unknown'
ORDER BY ComputerName;

DROP TABLE #ClientCompliance;
"@
}

#endregion

#region Import modules

$requiredModules = @(
    "dbatools",
    "Send-MailKitMessage"
)

Write-Log "Checking required PowerShell modules."
Write-Log "PowerShell executable: $($PSHome)\powershell.exe"
Write-Log "PowerShell edition: $($PSVersionTable.PSEdition)"
Write-Log "PowerShell version: $($PSVersionTable.PSVersion)"
Write-Log "User running task: $env:USERDOMAIN\$env:USERNAME"
Write-Log "PSModulePath: $env:PSModulePath"

foreach ($moduleName in $requiredModules) {
    Write-Log "Starting check for module '$moduleName'."

    try {
        $availableModules = @(
            Get-Module -ListAvailable -Name $moduleName -ErrorAction Stop
        )

        if ($availableModules.Count -eq 0) {
            Write-Log -Level "ERROR" -Message "Module '$moduleName' was not found for this user."

            throw @"
Required module '$moduleName' is not installed or is not available
to the scheduled-task account '$env:USERDOMAIN\$env:USERNAME'.

Install the module while logged in as that account:

Install-Module $moduleName -Scope CurrentUser -Force
"@
        }

        $selectedModule = $availableModules |
            Sort-Object Version -Descending |
            Select-Object -First 1

        Write-Log @"
Module '$moduleName' found.
Version: $($selectedModule.Version)
Path: $($selectedModule.Path)
"@

        Write-Log "Importing module '$moduleName'."

        Import-Module -Name $selectedModule.Path -Global -Force -ErrorAction Stop

        Write-Log "Module '$moduleName' imported successfully."
    }
    catch {
        Write-Log -Level "ERROR" -Message "Module '$moduleName' failed: $($_.Exception.Message)"
        throw
    }
}

Write-Log "All required PowerShell modules are available."

#endregion

#region Query database for each scheduled collection

Write-Log "Starting database queries for $($CollectionsToRun.Count) scheduled collection(s)."

Set-DbatoolsInsecureConnection -SessionOnly

$CollectionResults = @(
    foreach ($collection in $CollectionsToRun) {
        Write-Log "Querying collection '$($collection.Name)' against SUG '$($collection.Title)'."

        $query = Get-CollectionComplianceQuery `
            -CollectionName $collection.Name `
            -Title $collection.Title `
            -XDays $collection.XDays

        try {
            $queryStart = Get-Date

            $dataSet = Invoke-DbaQuery `
                -SqlInstance $SqlInstance `
                -Database $Database `
                -Query $query `
                -As DataSet `
                -EnableException

            $queryDuration = (Get-Date) - $queryStart

            $summaryTable      = $dataSet.Tables[0]
            $nonCompliantTable = $dataSet.Tables[1]
            $unknownTable      = $dataSet.Tables[2]

            $summaryRow = $summaryTable | Select-Object -First 1

            $totalCount        = [int]$summaryRow.Total
            $compliantCount    = [int]$summaryRow.Compliant
            $nonCompliantCount = [int]$summaryRow.NonCompliant
            $unknownCount      = [int]$summaryRow.UnknownClients

            $nonCompliantComputers = @(
                $nonCompliantTable | ForEach-Object { $_.ComputerName }
            )

            $unknownComputers = @(
                $unknownTable | ForEach-Object { $_.ComputerName }
            )

            Write-Log "Collection '$($collection.Name)': Total=$totalCount Compliant=$compliantCount NonCompliant=$nonCompliantCount Unknown=$unknownCount (duration $($queryDuration.TotalSeconds.ToString('0.00'))s)"

            [pscustomobject]@{
                CollectionName         = $collection.Name
                Title                  = $collection.Title
                XDays                  = $collection.XDays
                DaysAfterPatchTuesday  = $collection.DaysAfterPatchTuesday
                Total                  = $totalCount
                Compliant              = $compliantCount
                NonCompliant           = $nonCompliantCount
                Unknown                = $unknownCount
                NonCompliantComputers  = $nonCompliantComputers
                UnknownComputers       = $unknownComputers
                QuerySucceeded         = $true
                ErrorMessage           = ""
            }
        }
        catch {
            Write-Log "Query failed for collection '$($collection.Name)': $($_.Exception.Message)" "ERROR"

            [pscustomobject]@{
                CollectionName         = $collection.Name
                Title                  = $collection.Title
                XDays                  = $collection.XDays
                DaysAfterPatchTuesday  = $collection.DaysAfterPatchTuesday
                Total                  = 0
                Compliant              = 0
                NonCompliant           = 0
                Unknown                = 0
                NonCompliantComputers  = @()
                UnknownComputers       = @()
                QuerySucceeded         = $false
                ErrorMessage           = $_.Exception.Message
            }
        }
    }
)

Write-Log "All scheduled collection queries completed."

#endregion

#region Prepare shared email assets

$logoHtml = ConvertTo-LogoHtml -LogoValue $HtmlLogo -DefaultMimeType $LogoMimeType

#endregion

#region Build and send one email per collection

Write-Log "Starting email generation."

$GeneratedHtmlFiles = @()

foreach ($result in $CollectionResults) {
    $collectionNameHtml = ConvertTo-HtmlEncodedText -Value $result.CollectionName
    $titleHtml = ConvertTo-HtmlEncodedText -Value $result.Title
    $subjectForCollection = "{0} - {1} - {2}" -f $SubjectBase, $result.CollectionName, $result.DaysAfterPatchTuesday

    if (-not $result.QuerySucceeded) {
        $errorHtml = ConvertTo-HtmlEncodedText -Value $result.ErrorMessage

        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$subjectForCollection</title>
</head>
<body style="margin:0;padding:20px;background-color:#FFFFFF;">
    <div style="max-width:900px;margin:0 auto;background-color:#ffffff;padding:20px;border:1px solid #dddddd;">
        $logoHtml
        <h2 style="margin:0 0 10px 0;font-family:Arial,Helvetica,sans-serif;color:#1f4e78;">
            $subjectForCollection
        </h2>
        <p style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#444444;">
            Query failed for collection: <strong>$collectionNameHtml</strong>
            <br />
            Patch Tuesday offset: +$($result.DaysAfterPatchTuesday) day(s)
            <br />
            Error: <strong>$errorHtml</strong>
        </p>
    </div>
</body>
</html>
"@

        $htmlFilePath = Join-Path `
            -Path $env:TEMP `
            -ChildPath "ServerComplianceReport_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$($result.CollectionName -replace '[^a-zA-Z0-9]','_').html"

        if ($KeepHtmlFile) {
            Set-Content -LiteralPath $htmlFilePath -Value $emailBody -Encoding UTF8
            $GeneratedHtmlFiles += $htmlFilePath
            Write-Log "HTML body for '$($result.CollectionName)' saved to: $htmlFilePath"
        }

        $mailParams = @{
            SMTPServer                     = $SmtpServer
            Port                           = $SmtpPort
            From                           = $From
            RecipientList                  = $To
            Subject                        = $subjectForCollection
            HTMLBody                       = $emailBody
            UseSecureConnectionIfAvailable = $false
        }

        try {
            Send-MailKitMessage @mailParams
            Write-Log "Email sent successfully for collection '$($result.CollectionName)'."
        }
        catch {
            Write-Log "Email transmission failed for collection '$($result.CollectionName)': $($_.Exception.Message)" "ERROR"
        }

        continue
    }

    $compliantPercent = if ($result.Total -gt 0) {
        [math]::Round(($result.Compliant / $result.Total) * 100, 1)
    }
    else {
        0
    }

    $summaryHtml = @"
<table
    role="presentation"
    cellspacing="0"
    cellpadding="0"
    border="0"
    style="border-collapse:collapse;width:100%;max-width:600px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#222222;margin-bottom:12px;"
>
    <thead>
        <tr style="background-color:#1f4e78;color:#ffffff;">
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Metric</th>
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Count</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td style="border:1px solid #d9d9d9;padding:8px;">Total servers evaluated</td>
            <td style="border:1px solid #d9d9d9;padding:8px;">$($result.Total)</td>
        </tr>
        <tr>
            <td style="border:1px solid #d9d9d9;padding:8px;background-color:#90ee90;">Compliant</td>
            <td style="border:1px solid #d9d9d9;padding:8px;background-color:#90ee90;">$($result.Compliant) ($compliantPercent%)</td>
        </tr>
        <tr>
            <td style="border:1px solid #d9d9d9;padding:8px;background-color:#fa8072;font-weight:bold;">Non-Compliant</td>
            <td style="border:1px solid #d9d9d9;padding:8px;background-color:#fa8072;font-weight:bold;">$($result.NonCompliant)</td>
        </tr>
        <tr>
            <td style="border:1px solid #d9d9d9;padding:8px;background-color:#fffacd;">Unknown</td>
            <td style="border:1px solid #d9d9d9;padding:8px;background-color:#fffacd;">$($result.Unknown)</td>
        </tr>
    </tbody>
</table>
"@

    $nonCompliantHtml = New-ComputerListHtml `
        -ComputerNames $result.NonCompliantComputers `
        -EmptyMessage "No non-compliant servers were found." `
        -CellStyle "background-color:#fdecea;color:#000000;"

    $unknownHtml = New-ComputerListHtml `
        -ComputerNames $result.UnknownComputers `
        -EmptyMessage "No servers with unknown compliance state were found." `
        -CellStyle "background-color:#fffbea;color:#000000;"

    $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$subjectForCollection</title>
</head>
<body style="margin:0;padding:20px;background-color:#FFFFFF;">
    <div style="max-width:1000px;margin:0 auto;background-color:#ffffff;padding:20px;border:1px solid #dddddd;">
        $logoHtml

        <h2 style="margin:0 0 10px 0;font-family:Arial,Helvetica,sans-serif;color:#1f4e78;">
            $subjectForCollection
        </h2>

        <p style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#444444;">
            Collection: <strong>$collectionNameHtml</strong>
            <br />
            Software Update Group: <strong>$titleHtml</strong>
            <br />
            Patch Tuesday offset: <strong>+$($result.DaysAfterPatchTuesday) day(s)</strong>
            <br />
            Report date: <strong>$($AsOfDate.ToString("yyyy-MM-dd"))</strong>
        </p>

        <h3 style="font-family:Arial,Helvetica,sans-serif;color:#1f4e78;margin:20px 0 10px 0;">
            Summary
        </h3>

        $summaryHtml

        <h3 style="font-family:Arial,Helvetica,sans-serif;color:#1f4e78;margin:20px 0 10px 0;">
            Non-Compliant servers ($($result.NonCompliant))
        </h3>

        $nonCompliantHtml

        <h3 style="font-family:Arial,Helvetica,sans-serif;color:#1f4e78;margin:20px 0 10px 0;">
            Unknown compliance servers ($($result.Unknown))
        </h3>

        $unknownHtml
    </div>
</body>
</html>
"@

    $htmlFilePath = Join-Path `
        -Path $env:TEMP `
        -ChildPath "ServerComplianceReport_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$($result.CollectionName -replace '[^a-zA-Z0-9]','_').html"

    if ($KeepHtmlFile) {
        Set-Content -LiteralPath $htmlFilePath -Value $emailBody -Encoding UTF8
        $GeneratedHtmlFiles += $htmlFilePath
        Write-Log "HTML body for '$($result.CollectionName)' saved to: $htmlFilePath"
    }

    $mailParams = @{
        SMTPServer                     = $SmtpServer
        Port                           = $SmtpPort
        From                           = $From
        RecipientList                  = $To
        Subject                        = $subjectForCollection
        HTMLBody                       = $emailBody
        UseSecureConnectionIfAvailable = $false
    }

    try {
        Send-MailKitMessage @mailParams
        Write-Log "Email sent successfully for collection '$($result.CollectionName)'."
    }
    catch {
        Write-Log "Email transmission failed for collection '$($result.CollectionName)': $($_.Exception.Message)" "ERROR"
    }
}

Write-Log "Finished sending collection emails."

#endregion

#region Cleanup

if ($KeepHtmlFile) {
    Write-Log "KeepHtmlFile specified. HTML files retained in $env:TEMP."
}
else {
    foreach ($filePath in $GeneratedHtmlFiles) {
        if (Test-Path -LiteralPath $filePath) {
            Remove-Item -LiteralPath $filePath -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Log "Temporary HTML files removed."
}

#endregion

Write-Log "Server Compliance Report completed successfully."
Write-Log "============================================================"

Write-Output "Completed. Emails sent for $($CollectionResults.Count) collection(s)."
