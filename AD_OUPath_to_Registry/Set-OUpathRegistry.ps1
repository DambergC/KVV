$ErrorActionPreference = "SilentlyContinue"

# Get the computer's OU from Active Directory
$compName = $env:COMPUTERNAME
$root = [ADSI]""
$searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
$searcher.Filter = "(&(objectClass=computer)(cn=$compName))"
$computer = $searcher.FindOne()
$compliance = $false

# Registry path where we'll store the OU
$registryPath = "HKLM:\SOFTWARE\SCCMItems"
$registryName = "OUpath"

if ($computer) {
    $computerDN = $computer.Properties["distinguishedname"][0]
    # Extract OU path from distinguished name
    $ouPath = ($computerDN -split ',',2)[1]
    
    # Create the registry key if it doesn't exist
    if (!(Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    
    # Set the registry value
    Set-ItemProperty -Path $registryPath -Name $registryName -Value $ouPath -Force
    
    # Check if registry value was set correctly
    $currentValue = (Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue).$registryName
    
    if ($currentValue -eq $ouPath) {
        Write-Output "OU path successfully written to registry: $ouPath"
        $compliance = $true
    } else {
        Write-Output "Failed to write OU path to registry"
        $compliance = $false
    }
} else {
    Write-Output "Computer not found in AD"
    $compliance = $false
}

# Return compliance state for SCCM
return $compliance
