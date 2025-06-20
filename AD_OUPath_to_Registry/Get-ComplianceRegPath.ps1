# Check if registry value exists
$registryPath = "HKLM:\SOFTWARE\SCCMItems"
$registryName = "OUpath"

# Get the computer's current OU from AD for comparison
$compName = $env:COMPUTERNAME
$root = [ADSI]""
$searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
$searcher.Filter = "(&(objectClass=computer)(cn=$compName))"
$computer = $searcher.FindOne()

if ($computer) {
    $computerDN = $computer.Properties["distinguishedname"][0]
    $ouPathFromAD = ($computerDN -split ',',2)[1]
    
    # Check registry
    if (Test-Path $registryPath) {
        $ouPathInRegistry = (Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue).$registryName
        
        if ($ouPathInRegistry -eq $ouPathFromAD) {
            # Registry exists and is current
            return $true
        }
    }
}

# Either registry doesn't exist or is out of date
return $false
