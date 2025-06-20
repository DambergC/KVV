# Site configuration
$SiteCode = "KV1" # Site code
$ProviderMachineName = "vntsql0299.kvv.se" # SMS Provider machine name

# Customizations
$initParams = @{}
$initParams.Add("Verbose", $true) # Enable verbose logging
$initParams.Add("ErrorAction", "Stop") # Stop the script on any errors

# Error log file
$ErrorLogPath = "C:\temp\error_log.txt"

# Dynamic Throttling Configuration
$MaxThrottleLimit = 10 # Maximum allowed parallel tasks
$MinThrottleLimit = 2  # Minimum allowed parallel tasks

# Function to calculate dynamic throttle limit based on CPU usage
function Get-DynamicThrottleLimit {
    try {
        $cpuUsage = (Get-Counter -Counter "\Processor(_Total)\% Processor Time").CounterSamples.CookedValue
        if ($cpuUsage -lt 50) {
            return $MaxThrottleLimit # High ThrottleLimit when CPU usage is low
        } elseif ($cpuUsage -lt 80) {
            return [Math]::Max($MinThrottleLimit, [Math]::Floor($MaxThrottleLimit / 2)) # Medium ThrottleLimit
        } else {
            return $MinThrottleLimit # Low ThrottleLimit when CPU usage is high
        }
    } catch {
        Write-Error "Failed to calculate dynamic throttle limit: $_"
        $_ | Out-File -FilePath $ErrorLogPath -Append
        return $MinThrottleLimit # Default to minimum throttle limit in case of an error
    }
}

# Import the ConfigurationManager.psd1 module
try {
    if ((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams
    }
} catch {
    Write-Error "Failed to import ConfigurationManager module: $_"
    $_ | Out-File -FilePath $ErrorLogPath -Append
    exit 1
}

# Connect to the site's drive if it is not already present
try {
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }
    Set-Location "$($SiteCode):\" @initParams
} catch {
    Write-Error "Failed to connect to the Configuration Manager site drive: $_"
    $_ | Out-File -FilePath $ErrorLogPath -Append
    exit 1
}

# Data retrieval
try {
    Write-Host "Retrieving CM device objects from CM server by WMI" -ForegroundColor White
    $CMDevObjects = Get-WmiObject -ComputerName $ProviderMachineName -Class SMS_R_SYSTEM -Namespace root\sms\site_$SiteCode

    Write-Host "Retrieving DHCP scopes from DHCP server" -ForegroundColor White
    $dhcpScopes = Get-DhcpServerv4Scope -ComputerName "vntdhcp0002.kvv.se"

    Write-Host "Retrieving MITA computer objects from Active Directory" -ForegroundColor White
    $ADComputersMITA = Get-ADComputer -SearchBase "OU=Windows10 SMP,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se" -Filter {Name -like "*"} -Properties Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,Enabled,DistinguishedName

    Write-Host "Retrieving T1 computer objects from Active Directory" -ForegroundColor White
    $ADComputersT1 = Get-ADComputer -SearchBase "OU=Windows10 T1,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se" -Filter {Name -like "*"} -Properties Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,Enabled,DistinguishedName
} catch {
    Write-Error "Failed to retrieve data: $_"
    $_ | Out-File -FilePath $ErrorLogPath -Append
    exit 1
}

function Get-ScopeName($SubnetIP) {
    try {
        return ($dhcpScopes | Where-Object {$_.ScopeId -eq $SubnetIP}).Name
    } catch {
        Write-Error "Failed to retrieve scope name for subnet $SubnetIP: $_"
        $_ | Out-File -FilePath $ErrorLogPath -Append
    }
}

function Get-ComputerScopeNames($ComputerName) {
    try {
        $Subnets = ($CMDevObjects | Where-Object {$_.Name -eq $ComputerName}).IPSubnets
        $Subnets = $Subnets | Where-Object {$_ -notin ("10.0.0.0")}
        $SubnetNames += foreach ($Subnet in $Subnets) { Get-ScopeName $Subnet }
        return ($SubnetNames | Where-Object {$_ -ne $null}) -join ", "
    } catch {
        Write-Error "Failed to retrieve scope names for computer $ComputerName: $_"
        $_ | Out-File -FilePath $ErrorLogPath -Append
    }
}

$outputComputers = @()

# Parallel Script Block
$ParallelScriptBlock = {
    param ($ADComputer, $CMDevObjects, $dhcpScopes, $SiteCode, $ErrorLogPath)

    try {
        # Retrieve details for the current computer
        $SLSQLQuery = "SELECT SYS.Name0 AS Name,CS.Manufacturer0 AS Manufacturer,CS.Model0 AS Model FROM v_R_System AS SYS JOIN v_GS_COMPUTER_SYSTEM CS ON CS.ResourceID = SYS.ResourceID WHERE SYS.Name0 = '$($ADComputer.Name)'"
        $CMDevice = Get-CMDevice -Name $ADComputer.Name
        $SubnetScopes = Get-ComputerScopeNames -ComputerName $ADComputer.Name

        # Generate output object
        $ComputerDetails = Invoke-Sqlcmd -Database "CM_$SiteCode" -ServerInstance "vntsql0299.kvv.se" -Query $SLSQLQuery -TrustServerCertificate
        $myobj = [PSCustomObject]@{
            Datornamn              = $ADComputer.Name
            Plattform              = "Unknown" # Add logic to determine platform
            Serienummer            = $CMDevice.SerialNumber
            Operativsystem         = $ADComputer.OperatingSystem
            Operativsystemversion  = $ADComputer.OperatingSystemVersion
            Tillverkare            = $ComputerDetails.Manufacturer
            Modell                 = $ComputerDetails.Model
            Inloggningsdatum       = $ADComputer.LastLogonDate
            Aktiverad              = $ADComputer.Enabled
            DHCPScope              = $SubnetScopes
        }
        $myobj
    } catch {
        Write-Error "Error processing computer $($ADComputer.Name): $_"
        $_ | Out-File -FilePath $ErrorLogPath -Append
    }
}

# Process MITA and T1 computers with dynamic throttling
ForEach ($ADComputers in @($ADComputersMITA, $ADComputersT1)) {
    $ThrottleLimit = Get-DynamicThrottleLimit
    Write-Host "Processing with ThrottleLimit: $ThrottleLimit" -ForegroundColor Cyan

    $outputComputers += $ADComputers | ForEach-Object -Parallel $ParallelScriptBlock -ArgumentList $CMDevObjects, $dhcpScopes, $SiteCode, $ErrorLogPath -ThrottleLimit $ThrottleLimit
}

# Export the results
try {
    $outputComputers | Export-Csv -Path "C:\temp\computerslocation.csv" -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Host "Export completed successfully!" -ForegroundColor Green
} catch {
    Write-Error "Failed to export data: $_"
    $_ | Out-File -FilePath $ErrorLogPath -Append
}
