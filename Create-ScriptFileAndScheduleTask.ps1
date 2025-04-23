# Define the directory and script file
$ScriptDirectory = "C:\Program Files\BackupOutlookSign"
$ScriptFile = Join-Path -Path $ScriptDirectory -ChildPath "BackupRestore-OutlookSignatureFiles.ps1"

# Ensure the directory exists
if (-Not (Test-Path -Path $ScriptDirectory)) {
    New-Item -ItemType Directory -Path $ScriptDirectory -Force | Out-Null
    Write-Host "Created directory: $ScriptDirectory"
}

# Populate the script file with content from the GitHub repository
$ScriptContent = @"
<#
.SYNOPSIS
Backs up and restores Outlook Signature files with version control.

.DESCRIPTION
This script compares file timestamps between the local (source) and backup directories. It also restores files if no local files exist and stops execution if the backup folder path is unavailable.

.PARAMETER SourcePath
The path to the source Outlook Signature files. Defaults to the user's AppData folder.

.PARAMETER BackupPath
The path to the backup location. Defaults to a folder in the user's Documents directory.

.EXAMPLE
.\BackupRestore-OutlookSignatureFiles.ps1
#>

param (
    [string]$SourcePath = "$env:APPDATA\Microsoft\Signatures",
    [string]$BackupPath = "$env:HOMEDRIVE$env:HOMEPATH\Documents\OutlookSignaturesBackup"
)

# Enable strict mode
Set-StrictMode -Version Latest

# Function to write to the log
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$Timestamp - $Message"
}

# Ensure the backup folder exists
if (-Not (Test-Path -Path $BackupPath)) {
    Write-Log "Error: Backup folder path does not exist: $BackupPath"
    Write-Host "Script terminated as the backup folder is missing."
    exit 1
}

# Check if local files exist
if (-Not (Test-Path -Path $SourcePath) -or -Not (Get-ChildItem -Path $SourcePath -File)) {
    Write-Log "No local files found. Restoring files from the backup folder."

    # Ensure the source folder exists before restoring
    if (-Not (Test-Path -Path $SourcePath)) {
        New-Item -ItemType Directory -Path $SourcePath -Force | Out-Null
        Write-Log "Created source folder: $SourcePath"
    }

    # Restore files from backup to source
    Copy-Item -Path "$BackupPath\*" -Destination $SourcePath -Recurse -Force
    Write-Log "Restored files from $BackupPath to $SourcePath"
    Write-Host "Restoration complete."
} else {
    Write-Log "Local files exist. Proceeding with version control."

    # Version control: Compare file timestamps and sync
    Get-ChildItem -Path $SourcePath -Recurse | ForEach-Object {
        $SourceFile = $_
        $RelativePath = $SourceFile.FullName.Substring($SourcePath.Length).TrimStart('\')
        $BackupFile = Join-Path -Path $BackupPath -ChildPath $RelativePath

        # Handle directories
        if ($SourceFile.PSIsContainer) {
            if (-Not (Test-Path -Path $BackupFile)) {
                New-Item -ItemType Directory -Path $BackupFile -Force | Out-Null
                Write-Log "Created backup directory: $BackupFile"
            }
        } else {
            # Compare file timestamps
            if (Test-Path -Path $BackupFile) {
                $BackupFileInfo = Get-Item -Path $BackupFile
                if ($SourceFile.LastWriteTime -gt $BackupFileInfo.LastWriteTime) {
                    # Backup if the source file is newer
                    Copy-Item -Path $SourceFile.FullName -Destination $BackupFile -Force
                    Write-Log "Backed up: $($SourceFile.FullName) -> $BackupFile"
                } elseif ($SourceFile.LastWriteTime -lt $BackupFileInfo.LastWriteTime) {
                    # Restore if the backup file is newer
                    Copy-Item -Path $BackupFile -Destination $SourceFile.FullName -Force
                    Write-Log "Restored: $BackupFile -> $($SourceFile.FullName)"
                }
            } else {
                # Backup if the file doesn't exist in the backup directory
                Copy-Item -Path $SourceFile.FullName -Destination $BackupFile -Force
                Write-Log "Backed up (new file): $($SourceFile.FullName) -> $BackupFile"
            }
        }
    }
    Write-Log "Version control synchronization complete."
    Write-Host "Backup and restore process completed successfully."
}
"@

# Write the content to the file
Set-Content -Path $ScriptFile -Value $ScriptContent -Force
Write-Host "Populated script file: $ScriptFile"
