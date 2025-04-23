# Define the source and destination paths
$localPath = "$env:APPDATA\Microsoft\Signatures"
$backupPath = "\\server\backup\Signatures\$env:USERNAME"

# Function to compare file timestamps
function Compare-Timestamps {
    param (
        [string]$localFile,
        [string]$backupFile
    )
    $localTimestamp = (Get-Item $localFile).LastWriteTime
    $backupTimestamp = (Get-Item $backupFile).LastWriteTime

    if ($localTimestamp -gt $backupTimestamp) {
        return "LocalNewer"
    } elseif ($localTimestamp -lt $backupTimestamp) {
        return "BackupNewer"
    } else {
        return "Same"
    }
}

# Ensure the backup directory exists
if (!(Test-Path -Path $backupPath)) {
    New-Item -ItemType Directory -Path $backupPath
}

# Check if local signatures exist
if (Test-Path -Path $localPath) {
    # Check if backup signatures exist
    if (Test-Path -Path $backupPath) {
        # Compare files and decide whether to backup or restore
        $localFiles = Get-ChildItem -Path $localPath -Recurse
        foreach ($file in $localFiles) {
            $relativePath = $file.FullName.Substring($localPath.Length)
            $backupFile = Join-Path -Path $backupPath -ChildPath $relativePath

            if (Test-Path -Path $backupFile) {
                $comparison = Compare-Timestamps -localFile $file.FullName -backupFile $backupFile
                if ($comparison -eq "LocalNewer") {
                    Copy-Item -Path $file.FullName -Destination $backupFile -Force
                } elseif ($comparison -eq "BackupNewer") {
                    Copy-Item -Path $backupFile -Destination $file.FullName -Force
                }
            } else {
                Copy-Item -Path $file.FullName -Destination $backupFile -Force
            }
        }
    } else {
        # Backup local signatures if no backup exists
        Copy-Item -Path $localPath\* -Destination $backupPath -Recurse
    }
} else {
    # Restore from backup if no local signatures exist
    if (Test-Path -Path $backupPath) {
        Copy-Item -Path $backupPath\* -Destination $localPath -Recurse
    }
}
