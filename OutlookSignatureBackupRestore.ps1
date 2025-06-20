<#
.SYNOPSIS
    Script to back up and restore Outlook signatures for the currently logged-on user.

.DESCRIPTION
    This script handles the backup and restoration of Outlook signature files. It ensures that the backup directory exists
    and provides options for a dry run mode to preview actions without making changes.

.AUTHOR
    Christian Damberg

.CREATED
    2025-04-23

.PARAMETER LocalPath
    Specifies the local path where Outlook signatures are stored.
    Default: "$($env:APPDATA)\Microsoft\Signatures"

.PARAMETER BackupPath
    Specifies the path where the backup of the Outlook signatures will be stored.
    Default: "$($env:APPDATA)\Backup\Signatures"

.PARAMETER DryRun
    Switch to enable dry run mode, which simulates actions without making changes.

.EXAMPLES
    # Backup Outlook signatures to the default backup path
    .\OutlookSignatureBackupRestore.ps1

    # Restore Outlook signatures from a custom backup path
    .\OutlookSignatureBackupRestore.ps1 -BackupPath "C:\CustomBackupPath"

    # Run the script in dry run mode
    .\OutlookSignatureBackupRestore.ps1 -DryRun
#>

# Parameters for flexibility
param (
    [string]$LocalPath = "$($env:APPDATA)\Microsoft\Signatures",
    [string]$BackupPath = "$($env:APPDATA)\Backup\Signatures",
    [switch]$DryRun
)

# Log file path
$LogFilePath = Join-Path -Path $BackupPath -ChildPath "BackupRestore.log"

# Function to handle log rotation
function Rotate-Log {
    param (
        [string]$LogFilePath,
        [int]$MaxSizeKB = 1024, # Maximum log file size in KB (default: 1MB)
        [int]$MaxArchives = 5   # Maximum number of archived logs
    )

    try {
        if (Test-Path $LogFilePath) {
            # Check the size of the log file
            $logSizeKB = (Get-Item $LogFilePath).Length / 1KB
            if ($logSizeKB -ge $MaxSizeKB) {
                # Rotate logs
                for ($i = $MaxArchives - 1; $i -ge 1; $i--) {
                    $oldLog = "$LogFilePath.$i"
                    $newLog = "$LogFilePath.$($i + 1)"
                    if (Test-Path $oldLog) {
                        Rename-Item -Path $oldLog -NewName $newLog -Force
                    }
                }

                # Archive the current log
                Rename-Item -Path $LogFilePath -NewName "$LogFilePath.1" -Force

                # Create a new empty log file
                New-Item -Path $LogFilePath -ItemType File -Force | Out-Null
            }
        }
    } catch {
        Log-Message -Message "Log rotation failed: $_" -Level "Error"
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

    # Rotate logs before writing
    Rotate-Log -LogFilePath $LogFilePath -MaxSizeKB 1024 -MaxArchives 5

    # Write to log file
    Add-Content -Path $LogFilePath -Value $LogEntry

    # Display log entries based on level
    if ($Level -eq "Error") {
        Write-Host $LogEntry -ForegroundColor Red
    } elseif ($Level -eq "Warning") {
        Write-Host $LogEntry -ForegroundColor Yellow
    } else {
        Write-Host $LogEntry -ForegroundColor Green
    }
}

# Example usage of Log-Message in the script
Log-Message -Message "Script started." -Level "Info"

if ($DryRun) {
    Log-Message -Message "Dry run mode enabled. No changes will be made." -Level "Warning"
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
        Log-Message "Error comparing timestamps for ${localFile} and ${backupFile}: $_" -Level "Error"
    }
}

function Perform-Copy {
    param (
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$ActionDescription
    )
    if (-not $DryRun) {
        try {
            Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
            Log-Message "${ActionDescription}: ${SourcePath} -> ${DestinationPath}"
        } catch {
            Log-Message "Failed to ${ActionDescription}: $_" -Level "Error"
        }
    } else {
        Log-Message "Dry run: Would ${ActionDescription}: ${SourcePath} -> ${DestinationPath}"
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

    # Skip the log file
    if ($file.FullName -eq $LogFilePath) {
        Log-Message "Skipping log file during backup: $file"
        continue
    }

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

    # Skip the log file
    if ($backupFile.FullName -eq $LogFilePath) {
        Log-Message "Skipping log file during restore: $backupFile"
        continue
    }

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
