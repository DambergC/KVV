<#
.SYNOPSIS
    Collects device inventory from ConfigMgr and emails the results.

.DESCRIPTION
    This script extracts device inventory data from a ConfigMgr database,
    generates an HTML report, and emails the report to specified recipients.

.PARAMETER ConfigPath
    Path to the configuration XML file. Default is ".\ScriptConfigDevInvent.xml".

.PARAMETER NoEmail
    If specified, the script will not send emails.

.PARAMETER NoLog
    If specified, the script will not write to log files.

.PARAMETER GenerateCSV
    If specified, the script will also generate a CSV file with the inventory data.

.NOTES
    Version:        2.0
    Author:         Original by Christian Damberg
    Last Updated:   2025-05-19
    Updated by:     GitHub Copilot
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\ScriptConfigDevInvent.xml",
    
    [Parameter(Mandatory=$false)]
    [switch]$NoEmail,
    
    [Parameter(Mandatory=$false)]
    [switch]$NoLog,
    
    [Parameter(Mandatory=$false)]
    [switch]$GenerateCSV
)

#region Script Setup and Configuration
$scriptVersion = '2.0'
$scriptName = $MyInvocation.MyCommand.Name
$scriptStartTime = Get-Date

# Load configuration XML
try {
    [System.Xml.XmlDocument]$xml = Get-Content $ConfigPath -ErrorAction Stop
    Write-Verbose "Configuration file loaded successfully from $ConfigPath"
}
catch {
    Write-Error "Failed to load XML configuration file from $ConfigPath. Error: $_"
    exit 1
}

# Configuration settings
$siteServer = $xml.Configuration.SiteServer
$fileDate = Get-Date -Format yyyyMMdd
$HTMLFileSavePath = "G:\Scripts\Outfiles\DeviceInvent_$fileDate.HTML"
$CSVFileSavePath = "G:\Scripts\Outfiles\DeviceInvent_$fileDate.csv"
$SMTP = $xml.Configuration.MailSMTP
$MailFrom = $xml.Configuration.Mailfrom

# Handle mail port with default value and validation
$MailPortNumber = 25 # Default SMTP port
if (-not [string]::IsNullOrEmpty($xml.Configuration.MailPort)) {
    if ([int]::TryParse($xml.Configuration.MailPort, [ref]$null)) {
        $MailPortNumber = [int]::Parse($xml.Configuration.MailPort)
    }
}

$MailCustomer = $xml.Configuration.MailCustomer
$logFileName = $xml.Configuration.Logfile.Name
$logFilePath = $xml.Configuration.Logfile.Path
$logFile = Join-Path -Path $logFilePath -ChildPath $logFileName

# --- Early Log Write for Troubleshooting ---
try {
    $earlyLogMsg = "$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') [INFO] Script startup: $scriptName"
    Add-Content -Path $logFile -Value $earlyLogMsg
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') [INFO] Running as user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') [INFO] Working directory: $(Get-Location)"
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') [INFO] Log file path: $logFile"
} catch {
    Write-Error "Early log write failed: $_"
}

$today = Get-Date -Format yyyy-MM-dd
$monthName = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month)
$year = (Get-Date).Year
$titleDate = Get-Date -DisplayHint Date
$todayDefault = Get-Date -Format "yyyy-MM-dd"  # Define todayDefault for HTML footer
#endregion

#region Enhanced Logging Functions
function Write-Log {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$LogString,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFilePath = $script:logFile,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoTimestamp,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Severity = "INFO"
    )
    
    # Skip logging if NoLog switch is specified
    if ($NoLog) { 
        # Still output to console based on severity
        switch ($Severity) {
            "WARNING" { Write-Host $LogString -ForegroundColor Yellow }
            "ERROR" { Write-Host $LogString -ForegroundColor Red }
            default { Write-Verbose $LogString }
        }
        return 
    }
    
    try {
        # Create timestamp
        $timestamp = if (-not $NoTimestamp) { (Get-Date).ToString("yyyy/MM/dd HH:mm:ss") } else { "" }
        
        # Format message
        $logEntry = if (-not $NoTimestamp) {
            "$timestamp [$Severity] $LogString"
        } else {
            $LogString
        }
        
        # Add to log file
        Add-Content -Path $LogFilePath -Value $logEntry -ErrorAction Stop
        
        # Output to console with color based on severity
        switch ($Severity) {
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            default { Write-Verbose $logEntry }
        }
    }
    catch {
        Write-Error "Failed to write to log file $LogFilePath. Error: $_"
    }
}

function Rotate-Log {
    param(
        [string]$LogFilePath,
        [int]$ThresholdKB = 30,
        [int]$MaxLogFiles = 10
    )

    if (Test-Path -Path $LogFilePath) {
        $logFile = Get-Item $LogFilePath
        $logDir = $logFile.Directory.FullName
        $logBaseName = $logFile.BaseName
        $datetime = Get-Date -Format "yyyy-MM-dd-HHmm"

        # Check if log exceeds size threshold
        if ($logFile.Length/1KB -ge $ThresholdKB) {
            Write-Verbose "Log file '$($logFile.Name)' is larger than $ThresholdKB KB, rotating..."
            $newName = "$logDir\$logBaseName`_$datetime.log_old"
            Rename-Item -Path $LogFilePath -NewName $newName -Force
            
            # Create new empty log file
            New-Item -Path $LogFilePath -ItemType File | Out-Null
            Write-Log -LogString "Log file rotated, previous log saved as $newName" -NoTimestamp

            # Clean up old log files if there are too many
            $oldLogs = Get-ChildItem -Path $logDir -Filter "$logBaseName*old" | Sort-Object LastWriteTime
            if ($oldLogs.Count -gt $MaxLogFiles) {
                $logsToRemove = $oldLogs | Select-Object -First ($oldLogs.Count - $MaxLogFiles)
                $logsToRemove | ForEach-Object {
                    Write-Verbose "Removing old log file: $($_.FullName)"
                    Remove-Item -Path $_.FullName -Force
                }
            }
            
            return $true
        }
        return $false
    }
    else {
        # Create directory if it doesn't exist
        $logDirectory = Split-Path -Path $LogFilePath -Parent
        if (-not (Test-Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Create new log file
        New-Item -Path $LogFilePath -ItemType File -Force | Out-Null
        Write-Log -LogString "Log file created" -NoTimestamp
        return $false
    }
}

# Initialize logging
if (-not $NoLog) {
    # Create log directory if it doesn't exist
    $logDirectory = Split-Path -Path $logFile -Parent
    if (-not (Test-Path $logDirectory)) {
        try {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
            Write-Verbose "Created log directory: $logDirectory"
        }
        catch {
            Write-Error "Failed to create log directory $logDirectory. Error: $_"
            exit 1
        }
    }

    Write-Log -LogString "========================================" 
    Write-Log -LogString "$scriptName version $scriptVersion starting execution"
    Write-Log -LogString "Running as user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log -LogString "Running on computer: $env:COMPUTERNAME"
    Write-Log -LogString "========================================" 

    # Rotate log at start if needed
    Rotate-Log -LogFilePath $logFile -ThresholdKB 30 -MaxLogFiles 10
}
#endregion

#region Version Check and Validation
function Test-ScriptVersion {
    param(
        [string]$ConfigVersion,
        [string]$CurrentVersion = $scriptVersion
    )
    
    if (-not [string]::IsNullOrEmpty($ConfigVersion) -and $ConfigVersion -ne $CurrentVersion) {
        Write-Log -LogString "Script version mismatch! Config expects v$ConfigVersion but script is v$CurrentVersion" -Severity "WARNING"
        return $false
        write-host "False---"
    }
    return $true
}

# Check if expected version is specified in config
$expectedVersion = $xml.Configuration.ScriptVersion
if ($expectedVersion) {
    Test-ScriptVersion -ConfigVersion $expectedVersion
}

# Validate output paths
$outputFolder = Split-Path -Path $HTMLFileSavePath -Parent
if (-not (Test-Path $outputFolder)) {
    try {
        New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
        Write-Log -LogString "Created output directory: $outputFolder" -Severity "INFO"
    }
    catch {
        Write-Log -LogString "Failed to create output directory $outputFolder. Error: $_" -Severity "ERROR"
        exit 1
    }
}
#endregion

#region Module Management
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

# Required modules
$requiredModules = @("send-mailkitmessage", "PSWriteHTML", "PatchManagementSupportTools", "SQLPS")
$allModulesLoaded = $true

foreach ($module in $requiredModules) {
    $moduleLoaded = Import-RequiredModule -ModuleName $module
    $allModulesLoaded = $allModulesLoaded -and $moduleLoaded
}

if (-not $allModulesLoaded) {
    Write-Log -LogString "Not all required modules could be loaded. Exiting script." -Severity "ERROR"
    exit 1
}
#endregion

#region SQL Query and Data Collection
$query = @"
SELECT
    SYS.Netbios_Name0 AS [Device_Name],
    BIOS.SerialNumber0 AS [Serial_Number],
    CS.Model0 AS [Model],
    CS.Manufacturer0 AS [Manufacturer],
    OS.Caption0 AS [Operating_System],
    OS.Version0 AS [OS_Version],
    OS.BuildNumber0 AS [Build_Number],
    DeviceType.Device_Type,
    STRING_AGG(vru.Name0, ', ') AS [Primary_Users],
    SYS.User_Name0 AS [Last_Logon_User],
    SYS.Last_Logon_Timestamp0 AS [Last_Logon_Time],
    SYS.Resource_Domain_OR_Workgr0 AS [Domain],
    STRING_AGG(vru.Mail0, '; ') AS [User_Emails],
    IPADDR.IP_Addresses0 AS [IPv4_Address],
    bginfo.BoundaryName AS [Boundary_Name],
    bginfo.BoundaryValue AS [Boundary_Value],
    bginfo.BoundaryGroupName AS [Boundary_Group_Name],
    SCCM.Department00 AS [Department],
    SCCM.Extrapartition00 AS [Partition],
    SCCM.Jobtitle00 AS [Jobtitle],
    SCCM.Manager00 AS [Manager],
    SCCM.OUpath00 AS [OU_Path],
    CASE
        WHEN SCCM.OUpath00 LIKE '%Delad enhet%' THEN 'DMITA'
        WHEN SCCM.OUpath00 LIKE '%OU=Windows 11 MITAv2%' THEN 'MITA'
        WHEN SCCM.OUpath00 LIKE '%OU=MITA,OU=Fysiska klienter,OU=Windows 11%' THEN 'MITA'
        WHEN SCCM.OUpath00 LIKE '%OU=DMITA,OU=MITA,OU=Fysiska klienter%' THEN 'DMITA'
        WHEN SCCM.OUpath00 LIKE '%SMP%' THEN 'MITA'
        WHEN SCCM.OUpath00 LIKE '%T1%' THEN 'T1'
        ELSE '---'
    END AS [Classification],

    ISNULL(camtasia.DisplayName0, '---') AS [camtasia],
    ISNULL(bluebeam.DisplayName0, '---') AS [bluebeam],
    ISNULL(mindmanager.DisplayName0, '---') AS [mindmanager],
    ISNULL(msproject.DisplayName0, '---') AS [msproject],
    ISNULL(msvisio.DisplayName0, '---') AS [msvisio],
    ISNULL(stata.DisplayName0, '---') AS [stata],
    ISNULL(enterprisearchitect.DisplayName0, '---') AS [enterprisearchitect],
    ISNULL(philipsactiware.DisplayName0, '---') AS [philipsactiware],
    ISNULL(adobeacrobat.DisplayName0, '---') AS [adobeacrobat],
    ISNULL(assa.DisplayName0, '---') AS [assa],
    ISNULL(ibmspss.DisplayName0, '---') AS [ibmspss]

FROM
    v_R_System AS SYS
    JOIN v_GS_COMPUTER_SYSTEM AS CS ON SYS.ResourceID = CS.ResourceID
    JOIN v_GS_PC_BIOS AS BIOS ON SYS.ResourceID = BIOS.ResourceID
    JOIN v_GS_OPERATING_SYSTEM AS OS ON SYS.ResourceID = OS.ResourceID
    OUTER APPLY (
        SELECT
            CASE
                WHEN CS.Manufacturer0 LIKE '%VMware%' THEN 'Virtual'
                WHEN CS.Manufacturer0 LIKE '%Microsoft%' AND CS.Model0 LIKE '%Virtual%' THEN 'Virtual'
                WHEN CS.Manufacturer0 LIKE '%VirtualBox%' THEN 'Virtual'
                ELSE 'Physical'
            END AS Device_Type
    ) AS DeviceType
    LEFT JOIN v_UsersPrimaryMachines upm ON SYS.ResourceID = upm.MachineID
    LEFT JOIN v_R_User vru ON upm.UserResourceID = vru.ResourceID
    LEFT JOIN dbo.SCCMInventoryItems64_DATA SCCM ON SYS.ResourceID = SCCM.MachineID
    OUTER APPLY (
        SELECT TOP 1 IP.IP_Addresses0
        FROM v_RA_System_IPAddresses IP
        WHERE IP.ResourceID = SYS.ResourceID
          AND IP.IP_Addresses0 NOT LIKE '%:%'
          AND IP.IP_Addresses0 NOT LIKE '169.254.%'
        ORDER BY
            CASE
                WHEN IP.IP_Addresses0 LIKE '10.%' THEN 1
                WHEN IP.IP_Addresses0 LIKE '172.1[6-9].%' OR IP.IP_Addresses0 LIKE '172.2[0-9].%' OR IP.IP_Addresses0 LIKE '172.3[0-1].%' THEN 2
                WHEN IP.IP_Addresses0 LIKE '192.168.%' THEN 3
                ELSE 4
            END,
            IP.IP_Addresses0
    ) AS IPADDR
    OUTER APPLY (
        SELECT DISTINCT 
            bg.GroupID,
            bg.Name AS [BoundaryGroupName],
            b.DisplayName AS [BoundaryName],
            b.Value AS [BoundaryValue]
        FROM 
            vSMS_Boundary b
            JOIN vSMS_BoundaryGroupMembers bgm ON b.BoundaryID = bgm.BoundaryID
            JOIN vSMS_BoundaryGroup bg ON bgm.GroupID = bg.GroupID
        WHERE
            b.BoundaryType = 3 AND
            (CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 4)) * 16777216) +
            (CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 3)) * 65536) +
            (CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 2)) * 256) +
             CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 1))
            BETWEEN
            (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 4)) * 16777216) +
            (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 3)) * 65536) +
            (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 2)) * 256) +
             CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 1))
            AND
            (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 4)) * 16777216) +
            (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 3)) * 65536) +
            (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 2)) * 256) +
             CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 1))
    ) AS bginfo

    -- LEFT JOINs for software presence, now retrieving DisplayName0
    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%Camtasia%'
    ) AS camtasia ON SYS.ResourceID = camtasia.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%Bluebeam%'
    ) AS bluebeam ON SYS.ResourceID = bluebeam.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%Mindmanager%'
    ) AS mindmanager ON SYS.ResourceID = mindmanager.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%Microsoft Project%'
    ) AS msproject ON SYS.ResourceID = msproject.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%Microsoft Visio%'
    ) AS msvisio ON SYS.ResourceID = msvisio.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%stata%'
    ) AS stata ON SYS.ResourceID = stata.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS_64
        WHERE DisplayName0 LIKE '%Enterprise Architect%'
    ) AS enterprisearchitect ON SYS.ResourceID = enterprisearchitect.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS
        WHERE DisplayName0 LIKE '%Philips Actiware%'
    ) AS philipsactiware ON SYS.ResourceID = philipsactiware.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS
        WHERE DisplayName0 LIKE '%Adobe Acrobat DC%' OR DisplayName0 LIKE '%adobe acrobat 2%'
    ) AS adobeacrobat ON SYS.ResourceID = adobeacrobat.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS
        WHERE DisplayName0 LIKE '%assa per%'
    ) AS assa ON SYS.ResourceID = assa.ResourceID

    LEFT JOIN (
        SELECT ResourceID, DisplayName0
        FROM v_GS_ADD_REMOVE_PROGRAMS
        WHERE DisplayName0 LIKE '%IBM SPSS%'
    ) AS ibmspss ON SYS.ResourceID = ibmspss.ResourceID

WHERE
    DeviceType.Device_Type = 'Physical'
    AND (OS.Caption0 LIKE '%Windows 10%' OR OS.Caption0 LIKE '%Windows 11%')
    AND OS.Caption0 NOT LIKE '%Server%'
    AND SYS.Netbios_Name0 NOT LIKE 'WSDP%'
    AND (bginfo.BoundaryGroupName IS NULL OR bginfo.BoundaryGroupName NOT LIKE 'CL - Server Net%')
GROUP BY
    SYS.Netbios_Name0,
    BIOS.SerialNumber0,
    CS.Model0,
    CS.Manufacturer0,
    OS.Caption0,
    OS.Version0,
    OS.BuildNumber0,
    SYS.User_Name0,
    SCCM.Department00,
    SCCM.Jobtitle00,
    SCCM.Manager00,
    SYS.Last_Logon_Timestamp0,
    SYS.Resource_Domain_OR_Workgr0,
    IPADDR.IP_Addresses0,
    bginfo.BoundaryName,
    bginfo.BoundaryValue,
    bginfo.BoundaryGroupName,
    SCCM.OUpath00,
    DeviceType.Device_Type,
    SCCM.Extrapartition00,
    camtasia.DisplayName0,
    bluebeam.DisplayName0,
    mindmanager.DisplayName0,
    msproject.DisplayName0,
    msvisio.DisplayName0,
    stata.DisplayName0,
    enterprisearchitect.DisplayName0,
    philipsactiware.DisplayName0,
    adobeacrobat.DisplayName0,
    assa.DisplayName0,
    ibmspss.DisplayName0

ORDER BY
    SYS.Netbios_Name0;
"@

try {
    Write-Log -LogString "Running SQL query to fetch device inventory data..." -Severity "INFO"
    $data = Invoke-Sqlcmd -ServerInstance $siteServer -Database CM_KV1 -Query $query -ErrorAction Stop -QueryTimeout 600
    $deviceCount = $data | Measure-Object | Select-Object -ExpandProperty Count
    Write-Log -LogString "SQL query completed successfully, found $deviceCount devices" -Severity "INFO"
}
catch {
    Write-Log -LogString "Failed to execute SQL query: $_" -Severity "ERROR"
    Write-Log -LogString "Script execution terminated due to SQL error." -Severity "ERROR"
    exit 1
}

# Process the query results
$resultColl = @()
foreach ($row in $data) {
    $object = [PSCustomObject]@{
        'Device Name' = $row.device_name
        'Serial Number' = $row.Serial_number
        'Model' = $row.Model
        'Operating System' = $row.operating_system
        #'Build Number' = $row.Build_Number
        'Last Logon User' = $row.Last_logon_user
        'Department' = $row.Department
        #'Manager' = $row.Manager
        'Last Logon Time' = $row.Last_logon_Time
        'IPv4 Address' = $row.IPv4_Address
        #'Boundary Name' = $row.Boundary_Name
        'Boundary Group Name' = $row.Boundary_Group_Name
        'Type' = $row.Classification
        'Partition' = $row.partition
        'Camtasia' = $row.camtasia
        'Adobe' = $row.adobeacobat
        'BlueBeam' = $row.bluebeam
        'Mindmanager' = $row.mindmanager
        'Stata' = $row.stata
        'Enterprise Achitect' = $row.enterprisearchitect
        'Philips'= $row.philipsactiware
        'IBM Spss' = $row.ibmspss
        'MS Project' = $row.msproject
        'MS Visio' = $row.msvisio
        'Assa' = $row.assa
        'Primary User(s)' = $row.Primary_users

        
    }
    $resultColl += $object
}

Write-Log -LogString "Processed $($resultColl.Count) device records" -Severity "INFO"
#endregion

#region Report Generation
try {
    # Create HTML report
    Write-Log -LogString "Generating HTML report at $HTMLFileSavePath" -Severity "INFO"
    
    New-HTML -TitleText "Deviceinventering - Kriminalvården" -Online -FilePath $HTMLFileSavePath {
        
        New-HTMLHeader {
            New-HTMLSection -HeaderText 'Kriminalvården IT Arbetsplats' -HeaderTextSize 35 -BackgroundColor lightblue -HeaderTextColor Darkblue {
                New-HTMLPanel -BackgroundColor lightblue {
                    New-HTMLText -Text "Status Fysiska datorer $today från databasen för ConfigMgr" -FontSize 25 -Color Darkblue -FontFamily Arial -Alignment center -BackGroundColor lightblue
                    New-HTMLHorizontalLine
                }
            }
        }
        
        New-HTMLSection -Title "Sorterings- och exportbar data" -HeaderTextSize 20 -HeaderTextColor darkblue {
            New-HTMLTable -DataTable $resultColl -PagingLength 35 -AutoSize -Style nowrap
        }
        
        New-HTMLFooter {
            New-HTMLSection -Invisible {
                New-HTMLPanel -Invisible {
                    New-HTMLHorizontalLine
                    New-HTMLText -Text "Denna lista skapades $todayDefault" -FontSize 20 -Color Darkblue -FontFamily Arial -Alignment center -FontStyle italic
                }
            }
        }
    }
    
    Write-Log -LogString "HTML report generated successfully" -Severity "INFO"
    
    # Generate CSV if requested
    if ($GenerateCSV) {
        Write-Log -LogString "Generating CSV report at $CSVFileSavePath" -Severity "INFO"
        $resultColl | Export-Csv -Path $CSVFileSavePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Log -LogString "CSV report generated successfully" -Severity "INFO"
    }
}
catch {
    Write-Log -LogString "Failed to generate report: $_" -Severity "ERROR"
    exit 1
}
#endregion

#region Email Notification
if (-not $NoEmail) {
    Write-Log -LogString "Preparing email notification..." -Severity "INFO"
    
    # Get HTML content for email
    $content = Get-Content -Path $HTMLFileSavePath
    
    # Create HTML email body
    $Body = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Fysiska datorer ConfigMgr - Kriminalvården</title>
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
    H1 {
        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 18px;
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
</head>

<body>
    <p><h1>Inventering - fysiska klienter - Kriminalvården</h1></p> 
    <p>Bifogade filer innehåller dagsaktuell data from CM_KV1 databasen kopplad till ConfigMgr.<br><br>
<hr>
</p> 
    <p>Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
</body>
</html>
"@
    
    try {
        # Setup email parameters
        # Handle secure connection with default value
        $UseSecureConnectionIfAvailable = $false # Default value
        if (-not [string]::IsNullOrEmpty($xml.Configuration.UseSecureConnection)) {
            $secureConnValue = $xml.Configuration.UseSecureConnection.ToString().Trim().ToLower()
            if ($secureConnValue -eq "true" -or $secureConnValue -eq "1" -or $secureConnValue -eq "yes") {
                $UseSecureConnectionIfAvailable = $true
            }
        }
        
        # Handle authentication if configured
        $useAuth = $false
        if (-not [string]::IsNullOrEmpty($xml.Configuration.UseAuthentication)) {
            $authValue = $xml.Configuration.UseAuthentication.ToString().Trim().ToLower()
            if ($authValue -eq "true" -or $authValue -eq "1" -or $authValue -eq "yes") {
                $useAuth = $true
            }
        }
        
        # Handle credentials if authentication is enabled
        $Credential = $null
        if ($useAuth) {
            try {
                if ([string]::IsNullOrEmpty($xml.Configuration.EmailUsername) -or [string]::IsNullOrEmpty($xml.Configuration.EmailPassword)) {
                    Write-Log -LogString "Email authentication enabled but credentials missing in config" -Severity "WARNING"
                } else {
                    $securePassword = ConvertTo-SecureString $xml.Configuration.EmailPassword -AsPlainText -Force
                    $Credential = New-Object System.Management.Automation.PSCredential($xml.Configuration.EmailUsername, $securePassword)
                    Write-Log -LogString "Email authentication credentials loaded" -Severity "INFO"
                }
            }
            catch {
                Write-Log -LogString "Failed to set up email credentials: $_" -Severity "WARNING"
            }
        }
        
        # SMTP server settings
        $SMTPServer = $SMTP
        $Port = $MailPortNumber
        
        # From address
        $From = [MimeKit.MailboxAddress]$MailFrom
        
        # Recipient list - with improved error handling
        $RecipientList = [MimeKit.InternetAddressList]::new()
        try {
            # Handle different XML structures that might exist
            $recipientListXML = @()
            
            if ($xml.Configuration.Recipients.Recipients) {
                # Structure with nested Recipients
                foreach ($recipient in $xml.Configuration.Recipients.Recipients) {
                    if ($recipient.Email) {
                        $recipientListXML += $recipient.Email
                    }
                }
            } 
            elseif ($xml.Configuration.Recipients) {
                # Simpler structure
                if ($xml.Configuration.Recipients -is [string]) {
                    # Single recipient as string
                    $recipientListXML = @($xml.Configuration.Recipients)
                }
                elseif ($xml.Configuration.Recipients.Email) {
                    # Recipients with Email property
                    if ($xml.Configuration.Recipients.Email -is [array]) {
                        $recipientListXML = $xml.Configuration.Recipients.Email
                    }
                    else {
                        $recipientListXML = @($xml.Configuration.Recipients.Email)
                    }
                }
                else {
                    # Recipients collection
                    foreach ($recipient in $xml.Configuration.Recipients) {
                        if ($recipient.Email) {
                            $recipientListXML += $recipient.Email
                        }
                    }
                }
            }
            
            # Now add valid recipients to the list
            if ($recipientListXML -and $recipientListXML.Count -gt 0) {
                foreach ($Recipient in $recipientListXML) {
                    if (-not [string]::IsNullOrEmpty($Recipient)) {
                        try {
                            $RecipientList.Add([MimeKit.InternetAddress]$Recipient)
                            Write-Log -LogString "Added recipient: $Recipient" -Severity "INFO"
                        }
                        catch {
                            Write-Log -LogString "Invalid email address format for recipient: $Recipient" -Severity "WARNING"
                        }
                    }
                }
            }
            
            if ($RecipientList.Count -eq 0) {
                Write-Log -LogString "No valid email recipients found, email will not be sent" -Severity "WARNING"
                $NoEmail = $true
            }
        }
        catch {
            Write-Log -LogString "Error processing recipient list: $_" -Severity "ERROR"
            $NoEmail = $true
        }
        
        # BCC list with improved handling
        $BCCList = [MimeKit.InternetAddressList]::new()
        try {
            $bccRecipients = @()
            
            if ($xml.Configuration.BCCRecipients.Email) {
                # BCCRecipients with Email property
                if ($xml.Configuration.BCCRecipients.Email -is [array]) {
                    $bccRecipients = $xml.Configuration.BCCRecipients.Email
                }
                else {
                    $bccRecipients = @($xml.Configuration.BCCRecipients.Email)
                }
            }
            elseif ($xml.Configuration.BCCRecipients) {
                # Simple BCCRecipients
                if ($xml.Configuration.BCCRecipients -is [string]) {
                    $bccRecipients = @($xml.Configuration.BCCRecipients)
                }
                else {
                    # Collection of BCCRecipients
                    foreach ($bcc in $xml.Configuration.BCCRecipients) {
                        if ($bcc.Email) {
                            $bccRecipients += $bcc.Email
                        }
                    }
                }
            }
            
            # Now add valid BCC recipients
            if ($bccRecipients -and $bccRecipients.Count -gt 0) {
                foreach ($bccRecipient in $bccRecipients) {
                    if (-not [string]::IsNullOrEmpty($bccRecipient)) {
                        try {
                            $BCCList.Add([MimeKit.InternetAddress]$bccRecipient)
                            Write-Log -LogString "Added BCC recipient: $bccRecipient" -Severity "INFO"
                        }
                        catch {
                            Write-Log -LogString "Invalid email address format for BCC recipient: $bccRecipient" -Severity "WARNING"
                        }
                    }
                }
            }
        }
        catch {
            Write-Log -LogString "Error processing BCC recipient list: $_" -Severity "WARNING"
        }
        
        # Email subject with robust handling
        $Subject = "Deviceinventering"
        if (-not [string]::IsNullOrEmpty($MailCustomer)) {
            $Subject += " $MailCustomer"
        }
        $Subject += " $monthName $year"
        
        # HTML body
        $HTMLBody = [string]$Body
        
        # Attachment list
        $AttachmentList = [System.Collections.Generic.List[string]]::new()
        $AttachmentList.Add($HTMLFileSavePath)
        
        # Add CSV if generated
        if ($GenerateCSV -and (Test-Path $CSVFileSavePath)) {
            $AttachmentList.Add($CSVFileSavePath)
        }
        
        # Prepare email parameters
        $Parameters = @{
            "UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable
            "SMTPServer" = $SMTPServer
            "Port" = $Port
            "From" = $From
            "RecipientList" = $RecipientList
            "Subject" = $Subject
            "HTMLBody" = $HTMLBody
            "AttachmentList" = $AttachmentList
        }
        
        # Add credential if available
        if ($Credential) {
            $Parameters["Credential"] = $Credential
        }
        
        # Add BCC if available
        if ($BCCList.Count -gt 0) {
            $Parameters["BCCList"] = $BCCList
        }
        
        # Send email with more validation
        if ($RecipientList.Count -gt 0 -and -not [string]::IsNullOrEmpty($SMTPServer)) {
            try {
                Write-Log -LogString "Sending email to $($RecipientList.Count) recipients via ($SMTPServer):$Port" -Severity "INFO"
                Send-MailKitMessage @Parameters
                Write-Log -LogString "Email sent successfully to recipients" -Severity "INFO"
            }
            catch {
                Write-Log -LogString "Failed to send email: $_" -Severity "ERROR"
            }
        }
        else {
            Write-Log -LogString "Email not sent - missing required parameters (SMTP server or recipients)" -Severity "WARNING"
        }
    }
    catch {
        Write-Log -LogString "Error in email preparation: $_" -Severity "ERROR"
    }
}
else {
    Write-Log -LogString "Email notification skipped due to NoEmail parameter" -Severity "INFO"
}
#endregion

#region Script Cleanup
Set-Location $PSScriptRoot
$scriptEndTime = Get-Date
$executionTime = ($scriptEndTime - $scriptStartTime).TotalSeconds

Write-Log -LogString "Script execution completed in $executionTime seconds" -Severity "INFO"
Write-Log -LogString "========================================" 

# Clean up any temporary objects or connections
# Remove variables that might contain sensitive information
Remove-Variable -Name "Credential" -ErrorAction SilentlyContinue
Remove-Variable -Name "securePassword" -ErrorAction SilentlyContinue
#endregion
