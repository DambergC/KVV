<# 
.SYNOPSIS
Script to export device details from Configuration Manager.

.DESCRIPTION
This script retrieves device details from a Configuration Manager environment and exports them to a CSV.

.PARAMETER SiteCode
Site code of the Configuration Manager.

.PARAMETER ProviderMachineName
Name of the SMS Provider machine.

.PARAMETER OutputPath
Path to save the output CSV file.

.EXAMPLE
.\NEWScript.ps1 -SiteCode "KV1" -ProviderMachineName "vntsql0299.kvv.se" -OutputPath "C:\temp\output.csv"
#>

param(
    [string]$SiteCode = "KV1",
    [string]$ProviderMachineName = "vntsql0299.kvv.se",
    [string]$OutputPath = "C:\temp\ConfigMgr_Computers.csv"
)

# Logging setup
$LogFile = "C:\temp\ScriptLog.txt"
Write-Output "Script started at $(Get-Date)" | Out-File -Append $LogFile

# Import the ConfigurationManager module with error handling
try {
    if ((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    }
} catch {
    Write-Error "Failed to import ConfigurationManager module. $_"
    exit 1
}

# Connect to the ConfigMgr site drive with error handling
try {
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
    }
} catch {
    Write-Error "Unable to connect to ConfigMgr site drive. $_"
    exit 1
}

Set-Location "$($SiteCode):\"

Write-Host "Retrieving CM device objects from ConfigMgr via WMI" -ForegroundColor White
Write-Output "Retrieving CM device objects from ConfigMgr via WMI" | Out-File -Append $LogFile

# Retrieve all device information
try {
    $CMDevices = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName SMS_R_System
} catch {
    Write-Error "Failed to retrieve CM device objects. $_"
    exit 1
}

# Prepare the output array
$outputComputers = @()

# Create jobs for parallel processing
$jobs = @()
foreach ($CMDevice in $CMDevices) {
    $jobs += Start-Job -ScriptBlock {
        param ($SiteCode, $ResourceID)
        try {
            # Retrieve additional information about the device
            $ComputerDetails = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_G_System_COMPUTER_SYSTEM WHERE ResourceID = '$ResourceID'"
            $OperatingSystem = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_G_System_OPERATING_SYSTEM WHERE ResourceID = '$ResourceID'"
            $BoundaryGroups = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_BoundaryGroupMembers WHERE ResourceID = '$ResourceID'"

            # Gather details
            [PSCustomObject]@{
                Datornamn            = $ComputerDetails.Name
                Plattform            = if ($ComputerDetails.Model -like "*Server*") { "Server" } else { "Workstation" }
                Serienummer          = $ComputerDetails.SerialNumber
                Operativsystem       = $OperatingSystem.Caption
                Operativsystemversion = $OperatingSystem.Version
                Tillverkare          = $ComputerDetails.Manufacturer
                Modell               = $ComputerDetails.Model
                Inloggningsdatum     = $OperatingSystem.LastBootUpTime
                Aktiverad            = $CMDevice.Client
                Primäranvändare      = $CMDevice.PrimaryUser
                Användare            = $CMDevice.LastLogonUserName
                Användartitel        = "Unknown"
                Användarplats        = "Unknown"
                DHCPScope            = "Unknown"
                Boundarygrupp        = ($BoundaryGroups | ForEach-Object { $_.Name }) -join ", "
            }
        } catch {
            Write-Warning "Unable to retrieve details for device: $($CMDevice.Name). $_"
        }
    } -ArgumentList $SiteCode, $CMDevice.ResourceID
}

# Wait for all jobs to complete
$jobs | ForEach-Object { $_ | Wait-Job }

# Collect results from jobs
foreach ($job in $jobs) {
    $result = Receive-Job -Job $job
    if ($result) {
        $outputComputers += $result
    }
}

# Clean up jobs
$jobs | ForEach-Object { Remove-Job -Job $_ }

# Export the results to a CSV file
try {
    $outputComputers | Export-Csv -Path $OutputPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Host "Data exported to $OutputPath" -ForegroundColor Green
    Write-Output "Data exported to $OutputPath" | Out-File -Append $LogFile
} catch {
    Write-Error "Failed to export data to CSV. $_"
    exit 1
}
