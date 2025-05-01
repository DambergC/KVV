<#
.SYNOPSIS
This script retrieves information about Active Directory computers and their related Configuration Manager device objects. 
It exports details such as primary users, current users, operating system, and DHCP scope names for computers 
in specified Active Directory search bases to a CSV file.

.DESCRIPTION
The script performs the following tasks:
1. Connects to the Configuration Manager site.
2. Retrieves DHCP scopes from a specified DHCP server.
3. Fetches details of computers from Active Directory based on given search bases.
4. Retrieves Configuration Manager device objects and processes them in conjunction with AD data.
5. Exports the processed information to a CSV file.

.NOTES
Author: Christian Damberg, Telia Cygate AB
Date: 2025-05-01
Version: 1.0

#>


# Site configuration
$SiteCode = "KV1" # Site code
$ProviderMachineName = "vntsql0299.kvv.se" # SMS Provider machine name
$DHCPServerName = "vntdhcp0002.kvv.se"
$ADSearchBaseMITA = "OU=Windows10 SMP,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
$ADSearchBaseT1 = "OU=Windows10 T1,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

# Import ConfigurationManager Module
if (-not (Get-Module -Name ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
}

# Connect to CMSite Drive if not present
if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName -ErrorAction Stop
}

# Set the current location to the site code
Set-Location "$($SiteCode):\" -ErrorAction Stop

# Retrieve DHCP scopes once
Write-Host "Retrieving DHCP scopes from DHCP server..." -ForegroundColor White
$dhcpScopes = Get-DhcpServerv4Scope -ComputerName $DHCPServerName

# Function to get scope names for subnets
function Get-ScopeName {
    param ([string]$SubnetIP)
    ($dhcpScopes | Where-Object { $_.ScopeId -eq $SubnetIP }).Name
}

# Function to get computer scope names
function Get-ComputerScopeNames {
    param ([string]$ComputerName, [array]$CMDevObjects)
    $Subnets = ($CMDevObjects | Where-Object { $_.Name -eq $ComputerName }).IPSubnets
    $Subnets = $Subnets | Where-Object { $_ -ne "10.0.0.0" }
    $SubnetNames = foreach ($Subnet in $Subnets) { Get-ScopeName $Subnet }
    ($SubnetNames | Where-Object { $_ -ne $null }) -join ", "
}

# Retrieve Active Directory computers (batch processing)
function Get-ADComputers {
    param ([string]$SearchBase)
    Get-ADComputer -SearchBase $SearchBase -Filter * -Properties Name, OperatingSystem, OperatingSystemVersion, LastLogonDate, Enabled, DistinguishedName
}

Write-Host "Retrieving Active Directory computers..." -ForegroundColor White
$ADComputersMITA = Get-ADComputers -SearchBase $ADSearchBaseMITA
$ADComputersT1 = Get-ADComputers -SearchBase $ADSearchBaseT1

# Retrieve CM Device Objects once
Write-Host "Retrieving CM device objects..." -ForegroundColor White
$CMDevObjects = Get-CMDevice | Where-Object { $_.Name -ne "" }

# Function to process a batch of computers
function Process-Computers {
    param ([array]$ADComputers, [string]$Platform)
    $OutputComputers = @()

    foreach ($ADComputer in $ADComputers) {
        $CMDevice = $CMDevObjects | Where-Object { $_.Name -eq $ADComputer.Name }
        $PrimaryUsers = if ($CMDevice.PrimaryUser) {
            $CMDevice.PrimaryUser.Split(",").ForEach({
                (Get-ADUser $_.ToUpper().Replace("KVV\", "") -Properties Name, Title).Name
            }) -join ", "
        }
        $CurrentUsers = if ($CMDevice.UserName) {
            $CMDevice.UserName.Split(",").ForEach({
                (Get-ADUser $_.ToUpper().Replace("KVV\", "") -Properties Name, Title).Name
            }) -join ", "
        }

        $myobj = [PSCustomObject]@{
            Name               = $ADComputer.Name
            Platform           = $Platform
            SerialNumber       = $CMDevice.SerialNumber
            OperatingSystem    = $ADComputer.OperatingSystem
            OSVersion          = $ADComputer.OperatingSystemVersion
            PrimaryUsers       = $PrimaryUsers
            CurrentUsers       = $CurrentUsers
            DHCP_Scope_Names   = Get-ComputerScopeNames -ComputerName $ADComputer.Name -CMDevObjects $CMDevObjects
        }
        $OutputComputers += $myobj
    }

    return $OutputComputers
}

# Process MITA and T1 Computers in Parallel
$OutputComputers = @()
$OutputComputers += Process-Computers -ADComputers $ADComputersMITA -Platform "MITA"
$OutputComputers += Process-Computers -ADComputers $ADComputersT1 -Platform "T1"

# Export to CSV
Write-Host "Exporting results to CSV..." -ForegroundColor White
$OutputComputers | Export-Csv -Path C:\temp\computerslocation.csv -Delimiter ";" -NoTypeInformation -Encoding UTF8

Write-Host "Process completed!" -ForegroundColor Green
