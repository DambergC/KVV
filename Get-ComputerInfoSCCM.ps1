# Configuration Section
$SiteCode = "KV1"  # Site code (Update as per your environment)
$ProviderMachineName = "vntsql0299.kvv.se"  # SMS Provider machine name
$OutputPath = "C:\temp\ConfigMgr_Computers.csv"  # Path to save the output CSV file

# Import the ConfigurationManager module
if ((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
}

# Connect to the ConfigMgr site drive
if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
}

Set-Location "$($SiteCode):\"

Write-Host "Retrieving CM device objects from ConfigMgr via WMI" -ForegroundColor White

# Retrieve all device information
$CMDevices = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class SMS_R_System

# Prepare the output array
$outputComputers = @()

# Loop through each device in ConfigMgr
foreach ($CMDevice in $CMDevices) {
    Write-Host "Processing device: $($CMDevice.Name)" -ForegroundColor Yellow

    # Retrieve additional information about the device
    $ComputerDetails = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_G_System_COMPUTER_SYSTEM WHERE ResourceID = '$($CMDevice.ResourceID)'"
    $OperatingSystem = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_G_System_OPERATING_SYSTEM WHERE ResourceID = '$($CMDevice.ResourceID)'"
    $BoundaryGroups = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_BoundaryGroupMembers WHERE ResourceID = '$($CMDevice.ResourceID)'"

    # Gather details
    $Datornamn = $CMDevice.Name
    $Plattform = if ($ComputerDetails.Model -like "*Server*") { "Server" } else { "Workstation" }
    $Serienummer = $ComputerDetails.SerialNumber
    $Operativsystem = $OperatingSystem.Caption
    $Operativsystemversion = $OperatingSystem.Version
    $Tillverkare = $ComputerDetails.Manufacturer
    $Modell = $ComputerDetails.Model
    $Inloggningsdatum = $OperatingSystem.LastBootUpTime
    $Aktiverad = $CMDevice.Client
    $Primäranvändare = $CMDevice.PrimaryUser
    $Användare = $CMDevice.LastLogonUserName
    $Användartitel = "Unknown"  # ConfigMgr does not store user titles
    $Användarplats = "Unknown"  # ConfigMgr does not store user locations
    $DHCPScope = "Unknown"  # ConfigMgr does not store DHCP scopes
    $Boundarygrupp = ($BoundaryGroups | ForEach-Object { $_.Name }) -join ", "

    # Create a custom object
    $myobj = [PSCustomObject]@{
        Datornamn            = $Datornamn
        Plattform            = $Plattform
        Serienummer          = $Serienummer
        Operativsystem       = $Operativsystem
        Operativsystemversion = $Operativsystemversion
        Tillverkare          = $Tillverkare
        Modell               = $Modell
        Inloggningsdatum     = $Inloggningsdatum
        Aktiverad            = $Aktiverad
        Primäranvändare      = $Primäranvändare
        Användare            = $Användare
        Användartitel        = $Användartitel
        Användarplats        = $Användarplats
        DHCPScope            = $DHCPScope
        Boundarygrupp        = $Boundarygrupp
    }

    # Add the object to the output array
    $outputComputers += $myobj
}

# Export the results to a CSV file
$outputComputers | Export-Csv -Path $OutputPath -Delimiter ";" -NoTypeInformation -Encoding UTF8

Write-Host "Data exported to $OutputPath" -ForegroundColor Green
