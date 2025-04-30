# Load the required .NET types
Add-Type -TypeDefinition @"
using System;
using System.Collections.Concurrent;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
"@

# Configuration Section
$SiteCode = "KV1"
$ProviderMachineName = "vntsql0299.kvv.se"
$OutputPath = "C:\temp\ConfigMgr_Computers.csv"

# Retrieve all device information
$CMDevices = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName SMS_R_System

# Prepare the output collection
$Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()

# Create and configure the runspace pool
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, 10)  # Adjust max threads (10) as needed
$RunspacePool.Open()

# Create a list to track runspaces
$Runspaces = @()

# Start runspaces for each device
foreach ($CMDevice in $CMDevices) {
    $Runspace = [powershell]::Create().AddScript({
        param ($Device, $SiteCode)

        # Retrieve additional information about the device
        $ComputerDetails = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_G_System_COMPUTER_SYSTEM WHERE ResourceID = '$($Device.ResourceID)'"
        $OperatingSystem = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_G_System_OPERATING_SYSTEM WHERE ResourceID = '$($Device.ResourceID)'"

        # Create a custom object
        [PSCustomObject]@{
            Datornamn            = $Device.Name
            Plattform            = if ($ComputerDetails.Model -like "*Server*") { "Server" } else { "Workstation" }
            Serienummer          = $ComputerDetails.SerialNumber
            Operativsystem       = $OperatingSystem.Caption
            Operativsystemversion = $OperatingSystem.Version
            Tillverkare          = $ComputerDetails.Manufacturer
            Modell               = $ComputerDetails.Model
            Inloggningsdatum     = $OperatingSystem.LastBootUpTime
            Aktiverad            = $Device.Client
            Primäranvändare      = $Device.PrimaryUser
            Användare            = $Device.LastLogonUserName
        }
    }).AddArgument($CMDevice).AddArgument($SiteCode)

    $Runspace.RunspacePool = $RunspacePool
    $Runspaces += [PSCustomObject]@{ Pipe = $Runspace; Status = $Runspace.BeginInvoke(); }
}

# Wait for all runspaces to finish
foreach ($Runspace in $Runspaces) {
    $Runspace.Pipe.EndInvoke($Runspace.Status)
    $Results.Add($Runspace.Pipe.Invoke())
    $Runspace.Pipe.Dispose()
}

# Close the runspace pool
$RunspacePool.Close()
$RunspacePool.Dispose()

# Export results to a CSV file
$Results | Export-Csv -Path $OutputPath -Delimiter ";" -NoTypeInformation -Encoding UTF8

Write-Host "Data exported to $OutputPath" -ForegroundColor Green
