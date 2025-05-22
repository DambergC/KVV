$ErrorActionPreference = "SilentlyContinue"

# --- Section 1: Computer OU Info ---
$compName = $env:COMPUTERNAME
$root = [ADSI]""
$searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
$searcher.Filter = "(&(objectClass=computer)(cn=$compName))"
$computer = $searcher.FindOne()

# Registry path for SCCM Items
$registryPath = "HKLM:\SOFTWARE\SCCMInventoryItems"
$regComputerOU = "OUpath"

if ($computer) {
    $computerDN = $computer.Properties["distinguishedname"][0]
    $ouPath = ($computerDN -split ',',2)[1]

    if (!(Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $registryPath -Name $regComputerOU -Value $ouPath -Force
}

# --- Section 2: Logged-on User Info ---
# Get the currently logged-on user in DOMAIN\username format
$loggedOnUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
$regJobTitle   = "JobTitle"
$regDepartment = "Department"
$regManager    = "Manager"

if ($loggedOnUser) {
    $username = $loggedOnUser -replace '^.*\\', ''
    $userSearcher = New-Object System.DirectoryServices.DirectorySearcher($root)
    $userSearcher.Filter = "(&(objectClass=user)(sAMAccountName=$username))"
    $userObj = $userSearcher.FindOne()

    if ($userObj) {
        $jobTitle   = $userObj.Properties["title"][0]
        $department = $userObj.Properties["department"][0]
        $managerDN  = $userObj.Properties["manager"][0]

        # Extract manager CN (common name), not full DN
        if ($managerDN) {
            $managerCN = ($managerDN -split ',')[0] -replace '^CN=', ''
        } else {
            $managerCN = ""
        }

        # Create the registry key if it doesn't exist
        if (!(Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }

        # Set the registry values
        Set-ItemProperty -Path $registryPath -Name $regJobTitle   -Value ($jobTitle   | Out-String).Trim()   -Force
        Set-ItemProperty -Path $registryPath -Name $regDepartment -Value ($department | Out-String).Trim()   -Force
        Set-ItemProperty -Path $registryPath -Name $regManager    -Value ($managerCN  | Out-String).Trim()   -Force
    }
}

# --- Section 3: ExtraPartition ---
$regExtraPartition = "ExtraPartition"
$w10Path = "$env:ProgramFiles\Kriminalvarden\W10SO"
$w11Path = "$env:ProgramFiles\Kriminalvarden\W11SO"

if (Test-Path $w10Path) {
    $partitionValue = "W10SO"
} elseif (Test-Path $w11Path) {
    $partitionValue = "W11SO"
} else {
    $partitionValue = "None"
}

# Write ExtraPartition value to registry
if (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}
Set-ItemProperty -Path $registryPath -Name $regExtraPartition -Value $partitionValue -Force

# --- Section 4: Compliance Check ---
# Returns $true if all required properties are written, otherwise $false

# Check computer OU
$ouRegistryValue = (Get-ItemProperty -Path $registryPath -Name $regComputerOU -ErrorAction SilentlyContinue).$regComputerOU

# Check user info
$currentJobTitle   = (Get-ItemProperty -Path $registryPath -Name $regJobTitle   -ErrorAction SilentlyContinue).$regJobTitle
$currentDepartment = (Get-ItemProperty -Path $registryPath -Name $regDepartment -ErrorAction SilentlyContinue).$regDepartment
$currentManager    = (Get-ItemProperty -Path $registryPath -Name $regManager    -ErrorAction SilentlyContinue).$regManager

# Check ExtraPartition value
$currentExtraPartition = (Get-ItemProperty -Path $registryPath -Name $regExtraPartition -ErrorAction SilentlyContinue).$regExtraPartition

$compliance = $false
if ($ouRegistryValue -and $currentJobTitle -ne $null -and $currentDepartment -ne $null -and $currentManager -ne $null -and $currentExtraPartition -ne $null) {
    $compliance = $true
}
$compliance
