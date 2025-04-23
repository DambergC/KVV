<#
.SYNOPSIS
Backs up and optionally restores Outlook Signature files.

.DESCRIPTION
This script allows the user to back up Outlook Signature files and restore them when needed.
It includes error handling, parameterization, logging, and follows PowerShell best practices.

.PARAMETER SourcePath
The path to the source Outlook Signature files. Defaults to the user's AppData folder.

.PARAMETER BackupPath
The path to the backup location. Defaults to a folder in the user's Documents directory.

.PARAMETER Restore
Specifies whether to perform a restore instead of a backup.

.EXAMPLE
.\BackupRestore-OutlookSignatureFiles.ps1 -SourcePath "C:\Signatures" -BackupPath "D:\Backup"
.\BackupRestore-OutlookSignatureFiles.ps1 -Restore -BackupPath "D:\Backup"

.NOTES
Author: DambergC
Date: 2025-04-23
#>

param (
    [string]$SourcePath = "$env:APPDATA\Microsoft\Signatures",
    [string]$BackupPath = "$env:HOMEDRIVE$env:HOMEPATH\Documents\OutlookSignaturesBackup",
    [switch]$Restore
)

# Enable strict mode
Set-StrictMode -Version Latest

# Define a log file
$LogFile = Join-Path -Path $BackupPath -ChildPath "BackupRestoreLog.txt"

# Function to write to the log
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$Timestamp - $Message"
}

try {
    # Create the backup directory if it does not exist
    if (-Not (Test-Path -Path $BackupPath)) {
        New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
        Write-Log "Backup directory created at $BackupPath"
    }

    if ($Restore) {
        # Restore mode
        if (Test-Path -Path $BackupPath) {
            Copy-Item -Path "$BackupPath\*" -Destination $SourcePath -Recurse -Force
            Write-Log "Outlook Signatures restored from $BackupPath to $SourcePath"
            Write-Host "Outlook Signatures restored successfully from $BackupPath"
        } else {
            Write-Log "Restore failed: Backup directory does not exist at $BackupPath"
            Write-Host "Restore failed: No backup found at $BackupPath"
        }
    } else {
        # Backup mode
        if (Test-Path -Path $SourcePath) {
            Copy-Item -Path "$SourcePath\*" -Destination $BackupPath -Recurse -Force
            $FilesCopied = (Get-ChildItem -Path $SourcePath -Recurse).Count
            Write-Log "$FilesCopied files backed up from $SourcePath to $BackupPath"
            Write-Host "$FilesCopied Outlook Signature files have been backed up to $BackupPath"
        } else {
            Write-Log "Backup failed: Source directory does not exist at $SourcePath"
            Write-Host "Backup failed: No Outlook Signatures found at $SourcePath"
        }
    }
} catch {
    Write-Log "An error occurred: $_"
    Write-Host "An error occurred: $_"
}
