#Requires -RunAsAdministrator
<#
.SYNOPSIS
    TPS Deployment Script
.DESCRIPTION
    Creates TPS folder, cleans old files, copies new files from network share, and creates desktop shortcut
.PARAMETER TPSFolder
    Destination folder for TPS files (default: C:\TPS)
.PARAMETER NetworkShare
    Source network share path (default: \\SERVER\Share\TPS)
.PARAMETER SkipShortcut
    Skip desktop shortcut creation
#>

param(
    [string]$TPSFolder = "C:\TPS",
    [string]$NetworkShare = "\\SERVER\Share\TPS",
    [switch]$SkipShortcut
)

# Variables
$logFile = Join-Path $TPSFolder "TPS.log"
$publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
$shortcutPath = Join-Path $publicDesktop "TPS.lnk"

# Function to write log
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
    Write-Host $logMessage
}

# Start script
try {
    # Step 1: Create TPS folder if it doesn't exist
    if (-not (Test-Path $TPSFolder)) {
        New-Item -Path $TPSFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created folder: $TPSFolder"
    } else {
        Write-Log "Folder already exists: $TPSFolder"
    }

    # Ensure log file exists for subsequent logging
    if (-not (Test-Path $logFile)) {
        New-Item -Path $logFile -ItemType File -Force | Out-Null
    }

    # Step 2: Remove all files except *.log files
    Write-Log "Starting cleanup of $TPSFolder (excluding *.log files)"
    $filesToRemove = Get-ChildItem -Path $TPSFolder -File -Recurse | Where-Object { $_.Extension -ne ".log" }
    
    if ($filesToRemove) {
        foreach ($file in $filesToRemove) {
            try {
                Remove-Item $file.FullName -Force
                Write-Log "Removed file: $($file.FullName)"
            } catch {
                Write-Log "WARNING: Failed to remove $($file.FullName): $($_.Exception.Message)"
            }
        }
        Write-Log "Cleanup completed. Removed $($filesToRemove.Count) file(s)"
    } else {
        Write-Log "No files to remove"
    }

    # Step 3: Copy new files from network share
    if (Test-Path $NetworkShare) {
        Write-Log "Starting copy from network share: $NetworkShare"
        
        try {
            $copiedFiles = Copy-Item -Path "$NetworkShare\*" -Destination $TPSFolder -Recurse -PassThru -Force -ErrorAction Stop
            Write-Log "Copied $($copiedFiles.Count) item(s) from $NetworkShare"
            
            foreach ($file in $copiedFiles) {
                Write-Log "Copied: $($file.Name)"
            }
        } catch {
            Write-Log "ERROR: Failed to copy files: $($_.Exception.Message)"
            throw
        }
    } else {
        Write-Log "ERROR: Network share not found: $NetworkShare"
        throw "Network share not accessible"
    }

    # Step 4: Create shortcut on public desktop if it doesn't exist
    if (-not $SkipShortcut) {
        if (-not (Test-Path $shortcutPath)) {
            try {
                $WScriptShell = New-Object -ComObject WScript.Shell
                $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TPSFolder
                $shortcut.Description = "TPS Folder"
                $shortcut.Save()
                Write-Log "Created desktop shortcut: $shortcutPath"
                
                # Release COM object
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WScriptShell) | Out-Null
            } catch {
                Write-Log "WARNING: Failed to create shortcut: $($_.Exception.Message)"
            }
        } else {
            Write-Log "Desktop shortcut already exists: $shortcutPath"
        }
    } else {
        Write-Log "Skipped shortcut creation (SkipShortcut parameter set)"
    }

    Write-Log "Script completed successfully"
    exit 0

} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)"
    exit 1
}
