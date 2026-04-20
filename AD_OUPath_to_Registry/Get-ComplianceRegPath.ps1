param(
    [string]$RegistryPath = "HKLM:\SOFTWARE\SCCMItems",
    [string]$RegistryName = "OUpath"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertTo-LdapFilterSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $safeValue = $Value
    $safeValue = $safeValue -replace '\\', '\5c'
    $safeValue = $safeValue -replace '\*', '\2a'
    $safeValue = $safeValue -replace '\(', '\28'
    $safeValue = $safeValue -replace '\)', '\29'
    $safeValue = $safeValue -replace "`0", '\00'
    return $safeValue
}

$compliance = $false

try {
    $compName = $env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($compName)) {
        throw "Computer name was empty."
    }

    $safeCompName = ConvertTo-LdapFilterSafe -Value $compName
    $root = [ADSI]""
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
    $searcher.Filter = "(&(objectClass=computer)(cn=$safeCompName))"
    $computer = $searcher.FindOne()

    if (-not $computer) {
        Write-Output $compliance
        exit 1
    }

    $dnValues = $computer.Properties["distinguishedname"]
    if (-not $dnValues -or $dnValues.Count -eq 0) {
        throw "distinguishedName was missing for computer '$compName'."
    }

    $computerDN = [string]$dnValues[0]
    if ($computerDN -notmatch "^[^,]+,.+") {
        throw "Invalid distinguishedName format: '$computerDN'."
    }

    $ouPathFromAD = ($computerDN -split ',', 2)[1]

    if (-not (Test-Path -Path $RegistryPath)) {
        Write-Output $compliance
        exit 1
    }

    $regProperties = Get-ItemProperty -Path $RegistryPath -Name $RegistryName -ErrorAction Stop
    $ouPathInRegistry = [string]$regProperties.$RegistryName

    if ($ouPathInRegistry -eq $ouPathFromAD) {
        $compliance = $true
        Write-Output $compliance
        exit 0
    }

    Write-Output $compliance
    exit 1
}
catch {
    Write-Output $compliance
    exit 1
}
