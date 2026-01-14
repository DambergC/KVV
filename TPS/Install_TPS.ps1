<#
.SYNOPSIS
    TPS Deployment Script
.DESCRIPTION
    Creates TPS folder, cleans old files, copies new files from network share, and creates desktop shortcut
#>

# Variables
$tpsFolder = "C:\TPS"
$logFile = Join-Path $tpsFolder "TPS.log"
$networkShare = "\\SERVER\Share\TPS"  # TODO: Update this path
$publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
$shortcutPath = Join-Path $publicDesktop "TPS.lnk"

# Function to write log
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH: mm:ss"
    $logMessage = "$timestamp - $Message"
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage
}

# Start script
try {
    # Step 1: Create TPS folder if it doesn't exist
    if (-not (Test-Path $tpsFolder)) {
        New-Item -Path $tpsFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created folder: $tpsFolder"
    } else {
        Write-Log "Folder already exists: $tpsFolder"
    }

    # Ensure log file exists for subsequent logging
    if (-not (Test-Path $logFile)) {
        New-Item -Path $logFile -ItemType File -Force | Out-Null
    }

    # Step 2: Remove all files except *.log files
    Write-Log "Starting cleanup of $tpsFolder (excluding *.log files)"
    $filesToRemove = Get-ChildItem -Path $tpsFolder -File -Recurse | Where-Object { $_.Extension -ne ". log" }
    
    if ($filesToRemove) {
        foreach ($file in $filesToRemove) {
            Remove-Item $file.FullName -Force
            Write-Log "Removed file: $($file. FullName)"
        }
        Write-Log "Cleanup completed.  Removed $($filesToRemove.Count) file(s)"
    } else {
        Write-Log "No files to remove"
    }

    # Step 3: Copy new files from network share
    if (Test-Path $networkShare) {
        Write-Log "Starting copy from network share: $networkShare"
        $copiedFiles = Copy-Item -Path "$networkShare\*" -Destination $tpsFolder -Recurse -PassThru -Force
        Write-Log "Copied $($copiedFiles.Count) item(s) from $networkShare"
        
        foreach ($file in $copiedFiles) {
            Write-Log "Copied:  $($file.Name)"
        }
    } else {
        Write-Log "ERROR: Network share not found:  $networkShare"
        throw "Network share not accessible"
    }

    # Step 4: Create shortcut on public desktop if it doesn't exist
    if (-not (Test-Path $shortcutPath)) {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell. CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $tpsFolder
        $shortcut. Description = "TPS Folder"
        $shortcut.Save()
        Write-Log "Created desktop shortcut: $shortcutPath"
    } else {
        Write-Log "Desktop shortcut already exists: $shortcutPath"
    }

    Write-Log "Script completed successfully"
    exit 0

} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}
