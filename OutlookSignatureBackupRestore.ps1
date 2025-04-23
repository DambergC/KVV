# Parameters for flexibility
param (
    [string]$LocalPath = "$($env:APPDATA)\Microsoft\Signatures",
    [string]$BackupPath = "$($env:APPDATA)\Backup\Signatures"
)

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
        Write-Error "Error comparing timestamps for $localFile and $backupFile: $_"
    }
}

# Logging utility for better feedback
function Log-Message {
    param (
        [string]$message
    )
    Write-Host "$message"
}

# Ensure the backup directory exists
if (!(Test-Path -Path $BackupPath)) {
    try {
        New-Item -ItemType Directory -Path $BackupPath -Force
        Log-Message "Backup directory created at $BackupPath"
    } catch {
        Write-Error "Failed to create backup directory: $_"
        exit 1
    }
}

# Check if local signatures exist
if (Test-Path -Path $LocalPath) {
    # Check if backup signatures exist
    if (Test-Path -Path $BackupPath) {
        # Compare files and decide whether to backup or restore
        $localFiles = Get-ChildItem -Path $LocalPath -Recurse
        foreach ($file in $localFiles) {
            $relativePath = $file.FullName.Substring($LocalPath.Length).TrimStart('\')  # Ensure no leading backslash
            $backupFile = Join-Path -Path $BackupPath -ChildPath $relativePath

            if (Test-Path -Path $backupFile) {
                $comparison = Compare-Timestamps -localFile $file.FullName -backupFile $backupFile
                if ($comparison -eq "LocalNewer") {
                    try {
                        Copy-Item -Path $file.FullName -Destination $backupFile -Force
                        Log-Message "Updated backup for $file.FullName"
                    } catch {
                        Write-Error "Failed to update backup for $file.FullName: $_"
                    }
                } elseif ($comparison -eq "BackupNewer") {
                    try {
                        Copy-Item -Path $backupFile -Destination $file.FullName -Force
                        Log-Message "Restored $file.FullName from backup"
                    } catch {
                        Write-Error "Failed to restore $file.FullName from backup: $_"
                    }
                }
            } else {
                try {
                    Copy-Item -Path $file.FullName -Destination $backupFile -Force
                    Log-Message "Backed up $file.FullName"
                } catch {
                    Write-Error "Failed to back up $file.FullName: $_"
                }
            }
        }
    } else {
        # Backup local signatures if no backup exists
        try {
            Copy-Item -Path "$LocalPath\*" -Destination $BackupPath -Recurse
            Log-Message "All local signatures backed up to $BackupPath"
        } catch {
            Write-Error "Failed to back up local signatures: $_"
        }
    }
} else {
    # Restore from backup if no local signatures exist
    if (Test-Path -Path $BackupPath) {
        try {
            Copy-Item -Path "$BackupPath\*" -Destination $LocalPath -Recurse
            Log-Message "Restored all signatures from backup to $LocalPath"
        } catch {
            Write-Error "Failed to restore signatures from backup: $_"
        }
    } else {
        Write-Error "No local signatures or backup found. Nothing to do."
    }
}
