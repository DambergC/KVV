<#
.SYNOPSIS
    Get details for devices with Unknown deployment status in SCCM.

.DESCRIPTION
    Retrieves:
      - ComputerName
      - LastLogonUserName
      - LastClientActivity
      - ClientCheckStatus (Passed/Failed)
      - ClientActivityStatus (Active/Inactive)
    for devices that have "Unknown" status for a specified deployment.

.PARAMETER SiteServer
    The SCCM site server (FQDN or NetBIOS).

.PARAMETER SiteCode
    The SCCM site code (e.g. "P01").

.PARAMETER DeploymentID
    The SMS_DeploymentID of the deployment (e.g. "P0100123").

.PARAMETER ExportPath
    Optional path to export results to CSV.

.EXAMPLE
    .\Get-SCCMUnknownDeploymentStatus.ps1 -SiteServer "CM01.contoso.com" -SiteCode "P01" -DeploymentID "P0100123"

.EXAMPLE
    .\Get-SCCMUnknownDeploymentStatus.ps1 -SiteServer "CM01" -SiteCode "P01" -DeploymentID "P0100123" -ExportPath "C:\Temp\UnknownDevices.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteServer,

    [Parameter(Mandatory = $true)]
    [string]$SiteCode,

    [Parameter(Mandatory = $true)]
    [string]$DeploymentID,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath
)

# Build the WMI namespace for the site
$namespace = "root\SMS\site_$SiteCode"

Write-Verbose "Using namespace: $namespace on server: $SiteServer"

try {
    # 1. Get all deployment states for the specified deployment where status is Unknown.
    #    Unknown status is typically State = 0 in SMS_DeploymentSummary / DeploymentState.
    #    Here we use SMS_DeploymentState for per-device deployment state.
    $queryDeploymentState = "SELECT * FROM SMS_DeploymentState WHERE DeploymentID='$DeploymentID' AND State=0"

    Write-Verbose "Querying unknown deployment states with: $queryDeploymentState"

    $unknownStates = Get-WmiObject -Namespace $namespace -Class SMS_DeploymentState -ComputerName $SiteServer -Query $queryDeploymentState -ErrorAction Stop

    if (-not $unknownStates) {
        Write-Host "No devices with Unknown status found for DeploymentID $DeploymentID."
        return
    }

    # Get unique resource IDs
    $resourceIDs = $unknownStates | Select-Object -ExpandProperty ResourceID -Unique

    Write-Verbose "Found $($resourceIDs.Count) unique resource IDs with Unknown status."

    # 2. Pull device information from SMS_R_System and SMS_ClientSummary
    #    - SMS_R_System    -> Name, LastLogonUserName
    #    - SMS_ClientSummary -> ClientActiveStatus, ClientStateDescription (client check)
    #
    #    SMS_ClientSummary fields of interest:
    #      - ClientActiveStatus: 0=None, 1=Active, 2=Inactive
    #      - ClientState: Used for health; ClientStateDescription is a human description.
    #
    #    For last activity, SMS_ClientSummary has LastActiveTime.
    #

    # Build a comma-separated list of resource IDs for WMI queries
    $idList = ($resourceIDs | ForEach-Object { $_.ToString() }) -join ","

    $querySystem = "SELECT ResourceID, Name, LastLogonUserName FROM SMS_R_System WHERE ResourceID IN ($idList)"
    $querySummary = "SELECT ResourceID, ClientActiveStatus, ClientStateDescription, LastActiveTime FROM SMS_ClientSummary WHERE ResourceID IN ($idList)"

    Write-Verbose "Querying SMS_R_System with: $querySystem"
    Write-Verbose "Querying SMS_ClientSummary with: $querySummary"

    $systems = Get-WmiObject -Namespace $namespace -Query $querySystem -ComputerName $SiteServer -ErrorAction Stop
    $summaries = Get-WmiObject -Namespace $namespace -Query $querySummary -ComputerName $SiteServer -ErrorAction Stop

    # Index summaries by ResourceID for fast lookup
    $summaryHash = @{}
    foreach ($s in $summaries) {
        $summaryHash[$s.ResourceID] = $s
    }

    $results = foreach ($state in $unknownStates) {
        $resId = $state.ResourceID

        $sys = $systems   | Where-Object { $_.ResourceID -eq $resId }
        $sum = $summaryHash[$resId]

        if (-not $sys) { continue }

        # Map ClientActiveStatus to a text value
        $clientActivityStatus = switch ($sum.ClientActiveStatus) {
            1 { "Active" }
            2 { "Inactive" }
            default { "None/Unknown" }
        }

        # ClientStateDescription typically contains "Active", "Inactive", "Failed", "Unknown", etc.
        # We will classify Client Check as Passed vs Failed based on that text,
        # but still keep the original description.
        $clientCheckPassedOrFailed = if ($sum.ClientStateDescription -match "pass|healthy|good") {
            "Passed"
        } elseif ($sum.ClientStateDescription -match "fail|critical|error|inactive|unknown") {
            "Failed"
        } else {
            "Unknown"
        }

        # Last active time from summary can serve as "Last Activity"
        $lastActivity = $null
        if ($sum.LastActiveTime) {
            # Convert SCCM time format to DateTime
            $lastActivity = [System.Management.ManagementDateTimeConverter]::ToDateTime($sum.LastActiveTime)
        }

        [PSCustomObject]@{
            ResourceID              = $resId
            ComputerName            = $sys.Name
            LastLogonUser           = $sys.LastLogonUserName
            LastClientActivity      = $lastActivity
            ClientCheckStatus       = $clientCheckPassedOrFailed
            ClientCheckDescription  = $sum.ClientStateDescription
            ClientActivityStatus    = $clientActivityStatus
        }
    }

    if (-not $results -or $results.Count -eq 0) {
        Write-Host "No matching systems found in SMS_R_System / SMS_ClientSummary for the Unknown deployment states."
        return
    }

    # Show in console
    $results | Sort-Object ComputerName | Format-Table -AutoSize

    # Optionally export
    if ($ExportPath) {
        $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to $ExportPath"
    }

} catch {
    Write-Error "An error occurred: $_"
}