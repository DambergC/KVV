<#
.SYNOPSIS
Version controls and manages Outlook Signature files.

.DESCRIPTION
This script compares file timestamps between the source (local) and backup directories. 
It backs up files if the local version is newer and restores files if the backup version is newer.

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

# Function to compare and sync files
function Sync-Files {
    param (
        [string]$Source,
        [string]$Backup
    )

    # Ensure backup directory exists
    if (-Not (Test-Path -Path $Backup)) {
        New-Item -ItemType Directory -Path $Backup -Force | Out-Null
    }

    # Get all files from the source directory
    Get-ChildItem -Path $Source -Recurse | ForEach-Object {
        $SourceFile = $_
        $RelativePath = $SourceFile.FullName.Substring($Source.Length).TrimStart('\')
        $BackupFile = Join-Path -Path $Backup -ChildPath $RelativePath

        # Handle directories
        if ($SourceFile.PSIsContainer) {
            if (-Not (Test-Path -Path $BackupFile)) {
                New-Item -ItemType Directory -Path $BackupFile -Force | Out-Null
            }
        } else {
            # Compare file timestamps
            if (Test-Path -Path $BackupFile) {
                $BackupFileInfo = Get-Item -Path $BackupFile
                if ($SourceFile.LastWriteTime -gt $BackupFileInfo.LastWriteTime) {
                    # Backup if the source file is newer
                    Copy-Item -Path $SourceFile.FullName -Destination $BackupFile -Force
                    Write-Host "Backed up: $($SourceFile.FullName) -> $BackupFile"
                } elseif ($SourceFile.LastWriteTime -lt $BackupFileInfo.LastWriteTime) {
                    # Restore if the backup file is newer
                    Copy-Item -Path $BackupFile -Destination $SourceFile.FullName -Force
                    Write-Host "Restored: $BackupFile -> $($SourceFile.FullName)"
                }
            } else {
                # Backup if the file doesn't exist in the backup directory
                Copy-Item -Path $SourceFile.FullName -Destination $BackupFile -Force
                Write-Host "Backed up (new file): $($SourceFile.FullName) -> $BackupFile"
            }
        }
    }
}

# Perform the synchronization
Sync-Files -Source $SourcePath -Backup $BackupPath
