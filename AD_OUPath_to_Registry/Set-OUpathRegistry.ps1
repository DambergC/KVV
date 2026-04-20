param(
    [string]$RegistryPath = "HKLM:\SOFTWARE\SCCMItems",
    [string]$RegistryName = "OUpath"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "$timestamp [$Level] $Message"
}

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
        Write-Log -Level "ERROR" -Message "Computer '$compName' not found in AD."
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

    $ouPath = ($computerDN -split ',', 2)[1]

    if (-not (Test-Path -Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    New-ItemProperty -Path $RegistryPath -Name $RegistryName -PropertyType String -Value $ouPath -Force | Out-Null
    Set-ItemProperty -Path $RegistryPath -Name $RegistryName -Value $ouPath -Force

    $currentValue = (Get-ItemProperty -Path $RegistryPath -Name $RegistryName).$RegistryName
    if ($currentValue -eq $ouPath) {
        Write-Log -Level "INFO" -Message "OU path successfully written to registry: $ouPath"
        $compliance = $true
        Write-Output $compliance
        exit 0
    }

    Write-Log -Level "ERROR" -Message "Verification failed: registry value did not match expected OU path."
    Write-Output $compliance
    exit 1
}
catch {
    Write-Log -Level "ERROR" -Message "Script failed: $($_.Exception.Message)"
    Write-Output $compliance
    exit 1
}
