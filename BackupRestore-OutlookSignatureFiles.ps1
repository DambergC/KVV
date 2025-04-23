# Define the source path for Outlook Signatures
$sourcePath = "$env:APPDATA\Microsoft\Signatures"

# Define the destination path for backup
$backupPath = "$env:HOMEDRIVE$env:HOMEPATH\Documents\OutlookSignaturesBackup"

# Check if the source directory exists
if (Test-Path -Path $sourcePath) {
    # Ensure the backup directory exists
    if (-Not (Test-Path -Path $backupPath)) {
        New-Item -ItemType Directory -Path $backupPath | Out-Null
    }

    # Copy the contents of the Signatures folder to the backup location
    Copy-Item -Path $sourcePath\* -Destination $backupPath -Recurse -Force

    # Provide feedback to the user
    Write-Host "Outlook Signatures have been backed up to $backupPath"
} else {
    # Provide feedback if the source directory doesn't exist
    Write-Host "No Outlook Signatures found at $sourcePath"
}
