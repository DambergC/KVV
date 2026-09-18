<#
.SYNOPSIS
    Queries recent update bulletins and sends a static HTML email report.

.REQUIREMENTS
    Install-Module dbatools -Scope CurrentUser
    Install-Module Send-MailKitMessage -Scope CurrentUser

    Configuration file:
    D:\Scripts\UpdateBulletinReport\UpdateBulletinReport.config.xml
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "D:\Scripts\UpdateBulletinReport\UpdateBulletinReport.config.xml",

    [switch]$KeepHtmlFile
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

    # Console output is optional. It must never stop scheduled execution.
    try {
        switch ($Level) {
            "WARNING" {
                Write-Warning $Message
            }

            "ERROR" {
                Write-Error $Message -ErrorAction Continue
            }

            default {
                Write-Host $logLine
            }
        }
    }
    catch {
        # No interactive console may exist when running as a scheduled task.
    }

    # File logging is the authoritative log.
    try {
        $logDirectory = Split-Path -Path $LogPath -Parent

        if (
            -not [string]::IsNullOrWhiteSpace($logDirectory) -and
            -not (Test-Path -LiteralPath $logDirectory)
        ) {
            New-Item `
                -Path $logDirectory `
                -ItemType Directory `
                -Force `
                -ErrorAction Stop | Out-Null
        }

        Add-Content `
            -LiteralPath $LogPath `
            -Value $logLine `
            -Encoding UTF8 `
            -ErrorAction Stop
    }
    catch {
        # Do not allow a logging problem to terminate the report.
    }
}

#endregion

function Get-ConfigValue {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlDocument]$Xml,

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

#region Startup and configuration

Write-Log "============================================================"
Write-Log "Update Bulletin Report started."
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
    [xml]$configXml = Get-Content `
        -LiteralPath $ConfigPath `
        -Raw `
        -ErrorAction Stop

    Write-Log "XML configuration loaded successfully."
}
catch {
    Write-Log "Could not load XML configuration: $($_.Exception.Message)" "ERROR"
    throw
}

$SqlInstance = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Database/SqlInstance" `
    -Required

$Database = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Database/DatabaseName" `
    -Required

$SmtpServer = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Email/SmtpServer" `
    -Required

$SmtpPortText = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Email/SmtpPort" `
    -Required

$From = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Email/From" `
    -Required

$SubjectBase = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Email/Subject" `
    -Required

$HtmlLogo = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Email/HtmlLogo"

$LogoMimeType = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Email/HtmlLogoMimeType"

if ([string]::IsNullOrWhiteSpace($LogoMimeType)) {
    $LogoMimeType = "image/png"
}

$configuredLogPath = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Logging/LogPath"

if (-not [string]::IsNullOrWhiteSpace($configuredLogPath)) {
    $LogPath = $configuredLogPath
}

$DaysBackText = Get-ConfigValue `
    -Xml $configXml `
    -XPath "/Configuration/Query/DaysBack" `
    -Required

$recipientNodes = $configXml.SelectNodes(
    "/Configuration/Email/To/Recipient"
)

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
    $DaysBack = [int]$DaysBackText
}
catch {
    Write-Log "SMTP port or DaysBack is not numeric." "ERROR"
    throw
}

if ($SmtpPort -lt 1 -or $SmtpPort -gt 65535) {
    throw "SMTP port must be between 1 and 65535."
}

if ($DaysBack -lt 1) {
    throw "DaysBack must be greater than zero."
}

$ExcludeTitlePatterns = @(
    foreach ($node in $configXml.SelectNodes(
        "/Configuration/Query/ExcludeTitleContains/Pattern"
    )) {
        $value = ([string]$node.InnerText).Trim()

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $value
        }
    }
)

$ExcludeDescriptionPatterns = @(
    foreach ($node in $configXml.SelectNodes(
        "/Configuration/Query/ExcludeDescriptionContains/Pattern"
    )) {
        $value = ([string]$node.InnerText).Trim()

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $value
        }
    }
)

$ReportEndDate = Get-Date
$ReportStartDate = $ReportEndDate.AddDays(-$DaysBack)

$Subject = "{0} - {1} to {2}" -f `
    $SubjectBase,
    $ReportStartDate.ToString("yyyy-MM-dd"),
    $ReportEndDate.ToString("yyyy-MM-dd")

Write-Log "Log path: $LogPath"
Write-Log "SQL instance: $SqlInstance"
Write-Log "Database: $Database"
Write-Log "SMTP server: $SmtpServer"
Write-Log "SMTP port: $SmtpPort"
Write-Log "Recipient count: $($To.Count)"
Write-Log "DaysBack: $DaysBack"
Write-Log "Title exclusion count: $($ExcludeTitlePatterns.Count)"
Write-Log "Description exclusion count: $($ExcludeDescriptionPatterns.Count)"
Write-Log "Email subject: $Subject"

#endregion


# Fallback path used before XML configuration is loaded.
#$LogPath = "D:\Scripts\UpdateBulletinReport\LogFiles\UpdateBulletinReport.log"



#region Helper functions



function ConvertTo-HtmlEncodedText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if (
        $null -eq $Value -or
        [string]::IsNullOrWhiteSpace([string]$Value)
    ) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-EmailSafeLineBreaks {
    param(
        [AllowNull()]
        [object]$Value
    )

    $encodedText = ConvertTo-HtmlEncodedText -Value $Value

    return $encodedText -replace "(`r`n|`n|`r)", "<br />"
}

function Matches-Pattern {
    param(
        [AllowNull()]
        [string]$Text,

        [Parameter(Mandatory)]
        [string]$Pattern
    )

    if (
        [string]::IsNullOrWhiteSpace($Text) -or
        [string]::IsNullOrWhiteSpace($Pattern)
    ) {
        return $false
    }

    return (
        $Text.IndexOf(
            $Pattern,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -ge 0
    )
}

function Get-SeverityInfo {
    param(
        [AllowNull()]
        [object]$Severity
    )

    $severityText = if ($null -eq $Severity) {
        ""
    }
    else {
        ([string]$Severity).Trim().ToLowerInvariant()
    }

    switch ($severityText) {
        "10" {
            return @{
                Name        = "Critical"
                Cvss        = "9.0 - 10.0"
                Description = "Allows remote code execution or system takeover without user interaction. Fixes should be applied immediately."
                CellStyle   = "background-color:#fa8072;color:#000000;font-weight:bold;"
            }
        }

        "8" {
            return @{
                Name        = "Important"
                Cvss        = "7.0 - 8.9"
                Description = "Compromises data confidentiality, integrity, or user privileges. Fixes should be applied quickly."
                CellStyle   = "background-color:#ffa500;color:#000000;font-weight:bold;"
            }
        }

        "6" {
            return @{
                Name        = "Moderate"
                Cvss        = "4.0 - 6.9"
                Description = "Impacts non-default configurations or authentication requirements."
                CellStyle   = "background-color:#fffacd;color:#000000;"
            }
        }

        "2" {
            return @{
                Name        = "Low"
                Cvss        = "0.1 - 3.9"
                Description = "Minor component or characteristic adjustments with minimal security impact."
                CellStyle   = "background-color:#90ee90;color:#000000;"
            }
        }

        "0" {
            return @{
                Name        = "Unspecified"
                Cvss        = ""
                Description = "No severity classification was provided."
                CellStyle   = "background-color:#ffffff;color:#000000;"
            }
        }

        "critical" {
            return @{
                Name        = "Critical"
                Cvss        = "9.0 - 10.0"
                Description = "Allows remote code execution or system takeover without user interaction. Fixes should be applied immediately."
                CellStyle   = "background-color:#fa8072;color:#000000;font-weight:bold;"
            }
        }

        "important" {
            return @{
                Name        = "Important"
                Cvss        = "7.0 - 8.9"
                Description = "Compromises data confidentiality, integrity, or user privileges. Fixes should be applied quickly."
                CellStyle   = "background-color:#ffa500;color:#000000;font-weight:bold;"
            }
        }

        "moderate" {
            return @{
                Name        = "Moderate"
                Cvss        = "4.0 - 6.9"
                Description = "Impacts non-default configurations or authentication requirements."
                CellStyle   = "background-color:#fffacd;color:#000000;"
            }
        }

        "low" {
            return @{
                Name        = "Low"
                Cvss        = "0.1 - 3.9"
                Description = "Minor component or characteristic adjustments with minimal security impact."
                CellStyle   = "background-color:#90ee90;color:#000000;"
            }
        }

        default {
            return @{
                Name        = "Unspecified"
                Cvss        = ""
                Description = "No severity classification was provided. Database value: $severityText"
                CellStyle   = "background-color:#ffffff;color:#000000;"
            }
        }
    }
}

function ConvertTo-InfoUrlLink {
    param(
        [AllowNull()]
        [object]$Url
    )

    if (
        $null -eq $Url -or
        [string]::IsNullOrWhiteSpace([string]$Url)
    ) {
        return "<span style='color:#666666;'>No link available</span>"
    }

    try {
        $uri = [System.Uri]([string]$Url)

        if ($uri.Scheme -notin @("http", "https")) {
            throw "Only HTTP and HTTPS links are permitted."
        }

        $encodedUrl = [System.Net.WebUtility]::HtmlEncode(
            $uri.AbsoluteUri
        )

        return "<a href=""$encodedUrl"" style=""color:#0563C1;text-decoration:underline;"">Open article</a>"
    }
    catch {
        return "<span style='color:#b00020;'>Invalid URL</span>"
    }
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

    # Complete data URI.
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

    # Raw Base64 content.
    $base64Content = $LogoValue -replace "\s", ""

    if (
        $base64Content -match "^[A-Za-z0-9+/]*={0,2}$" -and
        ($base64Content.Length % 4 -eq 0)
    ) {
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

    # HTTP or HTTPS URL.
    try {
        $logoUri = [System.Uri]$LogoValue

        if ($logoUri.Scheme -in @("http", "https")) {
            $encodedLogoUrl = [System.Net.WebUtility]::HtmlEncode(
                $logoUri.AbsoluteUri
            )

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
        # Handled below.
    }

    Write-Log "Configured logo is not valid Base64 or HTTP/HTTPS." "WARNING"
    return ""
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
            Get-Module `
                -ListAvailable `
                -Name $moduleName `
                -ErrorAction Stop
        )

        if ($availableModules.Count -eq 0) {
            Write-Log `
                -Level "ERROR" `
                -Message "Module '$moduleName' was not found for this user."

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

        Import-Module `
            -Name $selectedModule.Path `
            -Global `
            -Force `
            -ErrorAction Stop

        Write-Log "Module '$moduleName' imported successfully."
    }
    catch {
        Write-Log `
            -Level "ERROR" `
            -Message "Module '$moduleName' failed: $($_.Exception.Message)"

        throw
    }
}

Write-Log "All required PowerShell modules are available."

#endregion

#region Query database

$query = @"
SELECT
    ui.DatePosted,
    ui.ArticleID,
    ui.Severity,
    ui.Title,
    ui.Description,
    ui.InfoURL
FROM v_UpdateInfo AS ui
WHERE ui.IsExpired = 0
  AND ui.IsSuperseded = 0
  AND ui.DatePosted >= DATEADD(DAY, -$DaysBack, GETDATE())
ORDER BY ui.DatePosted DESC;
"@

Write-Log "Starting database query."

try {
    Set-DbatoolsInsecureConnection -SessionOnly

    $queryStart = Get-Date

    $results = @(
        Invoke-DbaQuery `
            -SqlInstance $SqlInstance `
            -Database $Database `
            -Query $query `
            -EnableException
    )

    $queryDuration = (Get-Date) - $queryStart

    Write-Log "Database query completed."
    Write-Log "Raw result count: $($results.Count)"
    Write-Log "Query duration: $($queryDuration.TotalSeconds.ToString('0.00')) seconds"
}
catch {
    Write-Log "Database query failed: $($_.Exception.Message)" "ERROR"
    throw
}

#endregion

#region Apply exclusions and log matches

Write-Log "Starting title and description exclusion evaluation."

$filteredResults = @()
$excludedCount = 0
$includedCount = 0

foreach ($result in $results) {
    $matchedPatterns = @()

    $articleId = if ($null -eq $result.ArticleID) {
        ""
    }
    else {
        [string]$result.ArticleID
    }

    $title = if ($null -eq $result.Title) {
        ""
    }
    else {
        [string]$result.Title
    }

    $description = if ($null -eq $result.Description) {
        ""
    }
    else {
        [string]$result.Description
    }

    foreach ($pattern in $ExcludeTitlePatterns) {
        if (Matches-Pattern -Text $title -Pattern $pattern) {
            $matchedPatterns += "TITLE:'$pattern'"
        }
    }

    foreach ($pattern in $ExcludeDescriptionPatterns) {
        if (Matches-Pattern -Text $description -Pattern $pattern) {
            $matchedPatterns += "DESCRIPTION:'$pattern'"
        }
    }

    if ($matchedPatterns.Count -gt 0) {
        $excludedCount++

        Write-Log @"
EXCLUDED ArticleID=$articleId
Matched=$($matchedPatterns -join '; ')
Title=$title
"@
    }
    else {
        $filteredResults += $result
        $includedCount++

        Write-Log "KEPT ArticleID=$articleId"
    }
}

Write-Log "Exclusion evaluation completed."
Write-Log "Included count: $includedCount"
Write-Log "Excluded count: $excludedCount"

#endregion

#region Build HTML

Write-Log "Starting HTML generation."

$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm"

$htmlFilePath = Join-Path `
    -Path $env:TEMP `
    -ChildPath "RecentUpdates_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

$logoHtml = ConvertTo-LogoHtml `
    -LogoValue $HtmlLogo `
    -DefaultMimeType $LogoMimeType

$tableRows = [System.Text.StringBuilder]::new()

if ($filteredResults.Count -gt 0) {
    foreach ($result in $filteredResults) {
        $datePosted = if ($null -ne $result.DatePosted) {
            ([datetime]$result.DatePosted).ToString("yyyy-MM-dd HH:mm")
        }
        else {
            ""
        }

        $severityInfo = Get-SeverityInfo `
            -Severity $result.Severity

        $cellStyle = $severityInfo.CellStyle

        $datePostedHtml = ConvertTo-HtmlEncodedText `
            -Value $datePosted

        $articleIdHtml = ConvertTo-HtmlEncodedText `
            -Value $result.ArticleID

        $severityNameHtml = ConvertTo-HtmlEncodedText `
            -Value $severityInfo.Name

        $severityCvssHtml = ConvertTo-HtmlEncodedText `
            -Value $severityInfo.Cvss

        $severityDescriptionHtml = ConvertTo-HtmlEncodedText `
            -Value $severityInfo.Description

        $titleHtml = ConvertTo-HtmlEncodedText `
            -Value $result.Title

        $descriptionHtml = ConvertTo-EmailSafeLineBreaks `
            -Value $result.Description

        $infoUrlHtml = ConvertTo-InfoUrlLink `
            -Url $result.InfoURL

        [void]$tableRows.AppendLine(@"
<tr>
    <td style="border:1px solid #d9d9d9;padding:8px;vertical-align:top;white-space:nowrap;$cellStyle">
        $datePostedHtml
    </td>
    <td style="border:1px solid #d9d9d9;padding:8px;vertical-align:top;$cellStyle">
        $articleIdHtml
    </td>
    <td style="border:1px solid #d9d9d9;padding:8px;vertical-align:top;$cellStyle">
        <strong>$severityNameHtml</strong><br />
        <span style="font-size:11px;">CVSS: $severityCvssHtml</span><br />
        <span style="font-size:11px;font-weight:normal;">
            $severityDescriptionHtml
        </span>
    </td>
    <td style="border:1px solid #d9d9d9;padding:8px;vertical-align:top;$cellStyle">
        $titleHtml
    </td>
    <td style="border:1px solid #d9d9d9;padding:8px;vertical-align:top;$cellStyle">
        $descriptionHtml
    </td>
    <td style="border:1px solid #d9d9d9;padding:8px;vertical-align:top;white-space:nowrap;$cellStyle">
        $infoUrlHtml
    </td>
</tr>
"@)
    }

    $tableHtml = @"
<table
    role="presentation"
    cellspacing="0"
    cellpadding="0"
    border="0"
    style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222222;"
>
    <thead>
        <tr style="background-color:#1f4e78;color:#ffffff;">
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Date posted</th>
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Article ID</th>
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Severity</th>
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Title</th>
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Description</th>
            <th style="border:1px solid #d9d9d9;padding:8px;text-align:left;">Info URL</th>
        </tr>
    </thead>
    <tbody>
        $($tableRows.ToString())
    </tbody>
</table>
"@
}
else {
    $tableHtml = @"
<p style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222222;">
    No new bulletins were found in the last $DaysBack day(s).
</p>
"@
}

$htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$Subject</title>
</head>
<body style="margin:0;padding:20px;background-color:#FFFFFF;">
    <div style="max-width:1400px;margin:0 auto;background-color:#ffffff;padding:20px;border:1px solid #dddddd;">
        $logoHtml

        <h2 style="margin:0 0 10px 0;font-family:Arial,Helvetica,sans-serif;color:#1f4e78;">
            $Subject
        </h2>

        <p style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#444444;">
            Reporting period:
            <strong>$($ReportStartDate.ToString("yyyy-MM-dd"))</strong>
            to
            <strong>$($ReportEndDate.ToString("yyyy-MM-dd"))</strong>
            <br />
            Generated:
            <strong>$reportDate</strong>
            <br />
            Updates found:
            <strong>$($filteredResults.Count)</strong>
        </p>

        $tableHtml
    </div>
</body>
</html>
"@

Set-Content `
    -LiteralPath $htmlFilePath `
    -Value $htmlBody `
    -Encoding UTF8

Write-Log "HTML report created: $htmlFilePath"
Write-Log "HTML body length: $($htmlBody.Length) characters."

#endregion

#region Send email

$mailParams = @{
    SMTPServer                     = $SmtpServer
    Port                           = $SmtpPort
    From                           = $From
    RecipientList                  = $To
    Subject                        = $Subject
    HTMLBody                       = $htmlBody
    UseSecureConnectionIfAvailable = $false
}

Write-Log "Starting email transmission."
Write-Log "Recipient count: $($To.Count)"
Write-Log "Subject: $Subject"

try {
    Send-MailKitMessage @mailParams

    Write-Log "Email sent successfully."
}
catch {
    Write-Log "Email transmission failed: $($_.Exception.Message)" "ERROR"
    throw
}

#endregion

#region Cleanup

if ($KeepHtmlFile) {
    Write-Log "KeepHtmlFile specified. HTML file retained: $htmlFilePath"
}
else {
    Remove-Item `
        -LiteralPath $htmlFilePath `
        -Force `
        -ErrorAction SilentlyContinue

    Write-Log "Temporary HTML file removed."
}

#endregion

Write-Log "Update Bulletin Report completed successfully."
Write-Log "============================================================"

Write-Output "Report sent successfully. Query returned $($filteredResults.Count) included row(s)."