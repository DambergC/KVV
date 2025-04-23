# Parameters for flexibility
param (
    [string]$LocalPath = "$($env:APPDATA)\Microsoft\Signatures",
    [string]$BackupPath = "$($env:APPDATA)\Backup\Signatures",
    [switch]$DryRun
)

# Log file path
$LogFilePath = Join-Path -Path $BackupPath -ChildPath "BackupRestore.log"

# Ensure the backup directory exists
if (!(Test-Path -Path $BackupPath)) {
    try {
        New-Item -ItemType Directory -Path $BackupPath -Force
        Log-Message "Backup directory created at $BackupPath"
    } catch {
        Log-Message "Failed to create backup directory: $_" -Level "Error"
        exit 1
    }
}

# Logging utility for better feedback
function Log-Message {
    param (
        [string]$Message,
        [string]$Level = "Info" # Default to Info
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $LogFilePath -Value $LogEntry
    
    # Optional: Only write errors to the console
    if ($Level -eq "Error") {
        Write-Host $LogEntry -ForegroundColor Red
    } elseif ($Level -eq "Warning") {
        Write-Host $LogEntry -ForegroundColor Yellow
    } else {
        Write-Host $LogEntry
    }
}

if ($DryRun) {
    Log-Message "Dry run mode enabled. No changes will be made."
}

# Function to calculate the relative path
function Get-RelativePath {
    param (
        [string]$BasePath,
        [string]$TargetPath
    )
    $baseUri = New-Object System.Uri($BasePath)
    $targetUri = New-Object System.Uri($TargetPath)
    $relativeUri = $baseUri.MakeRelativeUri($targetUri)
    return $relativeUri.ToString().Replace('/', '\')
}

# Function to compare file timestamps
function Compare-Timestamps {
    param (
        [string]$localFile,
        [string]$backupFile
    )
    try {
        $localTimestamp = (Get-Item $localFile).LastWriteTime
        $backupTimestamp = (Get-Item $backupFile).LastWriteTime

        if ($localTimestamp -gt $backupTimestamp) {
            return "LocalNewer"
        } elseif ($localTimestamp -lt $backupTimestamp) {
            return "BackupNewer"
        } else {
            return "Same"
        }
    } catch {
        Log-Message "Error comparing timestamps for $localFile and $backupFile: $_" -Level "Error"
    }
}

# Utility function to handle file copy operations
function Perform-Copy {
    param (
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$ActionDescription
    )
    if (-not $DryRun) {
        try {
            Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
            Log-Message "$ActionDescription: $SourcePath -> $DestinationPath"
        } catch {
            Log-Message "Failed to $ActionDescription: $_" -Level "Error"
        }
    } else {
        Log-Message "Dry run: Would $ActionDescription: $SourcePath -> $DestinationPath"
    }
}

# Check if local signatures exist
if (Test-Path -Path $LocalPath) {
    # Check if backup signatures exist
    if (Test-Path -Path $BackupPath) {
        # Compare files and decide whether to backup or restore
        $localFiles = Get-ChildItem -Path $LocalPath -Recurse
        $backupFiles = Get-ChildItem -Path $BackupPath -Recurse

        # Process files in the local directory
        foreach ($file in $localFiles) {
            $relativePath = $file.FullName.Substring($LocalPath.Length).TrimStart('\')
            $backupFile = Join-Path -Path $BackupPath -ChildPath $relativePath

            if (Test-Path -Path $backupFile) {
                $comparison = Compare-Timestamps -localFile $file.FullName -backupFile $backupFile
                if ($comparison -eq "LocalNewer") {
                    Perform-Copy -SourcePath $file.FullName -DestinationPath $backupFile -ActionDescription "Update backup for"
                } elseif ($comparison -eq "BackupNewer") {
                    Perform-Copy -SourcePath $backupFile -DestinationPath $file.FullName -ActionDescription "Restore"
                }
            } else {
                Perform-Copy -SourcePath $file.FullName -DestinationPath $backupFile -ActionDescription "Backup"
            }
        }

        # Process files in the backup directory that are missing locally
        foreach ($backupFile in $backupFiles) {
            $relativePath = $backupFile.FullName.Substring($BackupPath.Length).TrimStart('\')
            $localFile = Join-Path -Path $LocalPath -ChildPath $relativePath

            if (-not (Test-Path -Path $localFile)) {
                Perform-Copy -SourcePath $backupFile.FullName -DestinationPath $localFile -ActionDescription "Restore missing file"
            }
        }
    } else {
        # Backup local signatures if no backup exists
        Perform-Copy -SourcePath "$LocalPath\*" -DestinationPath $BackupPath -ActionDescription "Backup all local signatures"
    }
} else {
    # Restore from backup if no local signatures exist
    if (Test-Path -Path $BackupPath) {
        Perform-Copy -SourcePath "$BackupPath\*" -DestinationPath $LocalPath -ActionDescription "Restore all signatures from backup"
    } else {
        Log-Message "No local signatures or backup found. Nothing to do." -Level "Error"
    }
}
