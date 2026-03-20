<#
.SYNOPSIS
  Cleans IIS log files (.log) older than a specified number of days.

.DESCRIPTION
  - Prompts for how many days of logs to keep (retention).
  - Finds log directories from the IIS configuration (site logFile.directory),
    expanding environment variables like %SystemDrive%.
  - Falls back to the default directory inetpub\logs\LogFiles if nothing is found.
  - Supports -WhatIf mode (preview).

.AUTHOR
  DambergC

.REQUIREMENTS
  - Running as Administrator is recommended.
  - Uses the WebAdministration module if available (normally on IIS servers).
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,3650)]
    [int]$KeepDays,

    [Parameter(Mandatory=$false)]
    [string[]]$ExtraPaths = @(),

    [switch]$IncludeFailedReqLogs # (optional) includes Failed Request Tracing logs (frt)
)

function Expand-Path {
    param([Parameter(Mandatory)][string]$Path)
    # Expand e.g. %SystemDrive%
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)

    # IIS can sometimes have relative paths; normalize carefully
    try { return (Resolve-Path -LiteralPath $expanded -ErrorAction Stop).Path }
    catch { return $expanded }
}

function Get-IISLogDirectories {
    $dirs = New-Object System.Collections.Generic.List[string]

    # Try via WebAdministration (IIS:\ provider)
    if (Get-Module -ListAvailable -Name WebAdministration) {
        Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null

        try {
            $sites = Get-ChildItem IIS:\Sites -ErrorAction Stop
            foreach ($s in $sites) {
                $d = $s.logFile.directory
                if ([string]::IsNullOrWhiteSpace($d)) { continue }
                $dirs.Add((Expand-Path $d))
            }
        } catch {
            # Ignore and fall back later
        }
    }

    # Default directory
    $default = Expand-Path "$env:SystemDrive\inetpub\logs\LogFiles"
    $dirs.Add($default)

    if ($IncludeFailedReqLogs) {
        $frt = Expand-Path "$env:SystemDrive\inetpub\logs\FailedReqLogFiles"
        $dirs.Add($frt)
    }

    foreach ($p in $ExtraPaths) {
        if (-not [string]::IsNullOrWhiteSpace($p)) {
            $dirs.Add((Expand-Path $p))
        }
    }

    # Unique + existing directories
    return @($dirs) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique |
        Where-Object { Test-Path -LiteralPath $_ }
}

# ====== Input ======
if (-not $PSBoundParameters.ContainsKey('KeepDays')) {
    $input = Read-Host "How many days of logs should be kept? (e.g. 30)"
    if (-not [int]::TryParse($input, [ref]$KeepDays)) {
        throw "Invalid value: '$input'. Please enter an integer (e.g. 30)."
    }
    if ($KeepDays -lt 1) { throw "KeepDays must be >= 1." }
}

$cutoff = (Get-Date).AddDays(-$KeepDays)

Write-Host "Keeping the most recent $KeepDays days of logs. Deleting .log files older than $($cutoff.ToString('yyyy-MM-dd HH:mm:ss'))."
Write-Host ""

# Force array so Length always exists (even if only one directory is returned)
$logDirs = @(Get-IISLogDirectories)

if ($logDirs.Length -eq 0) {
    throw "No log directories found to clean."
}

Write-Host "Log directories that will be searched:"
$logDirs | ForEach-Object { Write-Host " - $_" }
Write-Host ""

# ====== Deletion ======
$totalFiles = 0
$totalBytes = 0

foreach ($dir in $logDirs) {
    # IIS logs are typically *.log. If you also want *.zip etc, add them here.
    $files = Get-ChildItem -LiteralPath $dir -Recurse -File -Filter *.log -ErrorAction SilentlyContinue |
             Where-Object { $_.LastWriteTime -lt $cutoff }

    foreach ($f in $files) {
        $totalFiles++
        $totalBytes += $f.Length

        if ($PSCmdlet.ShouldProcess($f.FullName, "Remove-Item")) {
            try {
                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
            } catch {
                Write-Warning "Could not delete: $($f.FullName) :: $($_.Exception.Message)"
            }
        }
    }
}

$mb = [Math]::Round($totalBytes / 1MB, 2)
Write-Host ""
Write-Host "Done. Matched files: $totalFiles. Size: ~ $mb MB."
Write-Host "Tip: run with -WhatIf to preview deletions."
