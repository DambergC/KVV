#Requires -RunAsAdministrator
<#
.SYNOPSIS
    TPS Deployment Script - Copy files from source to C:\Program Files\TPS
    
.DESCRIPTION
    This script automates the deployment of TPS files:
    1. Moves current TPS files to backup folder
    2. Cleans backup folder if it contains old files
    3. Removes old files from TPS folder (excluding backup folder and log files)
    4. Copies new files from source
    5. Logs all operations to TPS folder
    
.PARAMETER SourcePath
    Source path where TPS files are located (network share or local path)
    
.PARAMETER DestinationPath
    Destination folder for TPS files (default: C:\Program Files\TPS)
    
.PARAMETER BackupFolderName
    Name of the backup folder (default: Backup)
    
.EXAMPLE
    .\Deploy-TPS.ps1 -SourcePath "\\SERVER\Share\TPS"
    
.EXAMPLE
    .\Deploy-TPS.ps1 -SourcePath "\\SERVER\Share\TPS" -DestinationPath "C:\Program Files\TPS"
    
.NOTES
    Author: DambergC
    Version: 1.0
    Date: 2026-01-26
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Source path for TPS files")]
    [ValidateScript({
        if (Test-Path $_ ) { $true }
        else { throw "Source path not found: $_" }
    })]
    [string]$SourcePath,
    
    [Parameter(HelpMessage="Destination folder for TPS files")]
    [string]$DestinationPath = "C:\Program Files\TPS",
    
    [Parameter(HelpMessage="Backup folder name")]
    [string]$BackupFolderName = "Backup"
)

# ============================================================================
# GLOBAL VARIABLES
# ============================================================================
$script:LogFile = $null
$script:BackupFolder = $null
$script:StartTime = Get-Date

# ============================================================================
# LOGGING FUNCTIONS
# ============================================================================
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        'INFO'    { Write-Host $logMessage -ForegroundColor White }
        'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
        'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
    }
    
    # Write to log file
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value $logMessage -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }
}

function Initialize-LogFile {
    param(
        [string]$LogPath
    )
    
    try {
        $logFileName = "TPS_Deployment_$(Get-Date -Format 'yyyyMMdd').log"
        $script:LogFile = Join-Path $LogPath $logFileName
        
        # Ensure directory exists
        if (-not (Test-Path $LogPath)) {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        }
        
        # Create or append to log file
        if (-not (Test-Path $script:LogFile)) {
            New-Item -Path $script:LogFile -ItemType File -Force | Out-Null
        }
        
        Write-Log "========================================" -Level INFO
        Write-Log "TPS Deployment Script Started" -Level INFO
        Write-Log "========================================" -Level INFO
        
        return $true
    }
    catch {
        Write-Warning "Failed to initialize log file: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# BACKUP FUNCTIONS
# ============================================================================
function Clear-BackupFolder {
    param(
        [string]$BackupPath
    )
    
    Write-Log "Checking backup folder: $BackupPath" -Level INFO
    
    if (Test-Path $BackupPath) {
        try {
            $backupFiles = Get-ChildItem -Path $BackupPath -Recurse -Force
            
            if ($backupFiles.Count -gt 0) {
                Write-Log "Backup folder contains $($backupFiles.Count) items - cleaning..." -Level WARNING
                
                # Remove all items from backup folder
                Get-ChildItem -Path $BackupPath -Recurse -Force | Remove-Item -Force -Recurse -ErrorAction Stop
                
                Write-Log "Backup folder cleaned successfully" -Level SUCCESS
            }
            else {
                Write-Log "Backup folder is empty - no cleanup needed" -Level INFO
            }
        }
        catch {
            Write-Log "Error cleaning backup folder: $($_.Exception.Message)" -Level ERROR
            throw
        }
    }
    else {
        Write-Log "Backup folder does not exist - will be created" -Level INFO
    }
}

function Move-FilesToBackup {
    param(
        [string]$SourceFolder,
        [string]$BackupPath
    )
    
    Write-Log "Moving TPS files to backup folder..." -Level INFO
    
    try {
        # Create backup folder if it doesn't exist
        if (-not (Test-Path $BackupPath)) {
            New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
            Write-Log "Created backup folder: $BackupPath" -Level INFO
        }
        
        # Get all items except backup folder and log files
        $itemsToBackup = Get-ChildItem -Path $SourceFolder -Force | Where-Object {
            $_.Name -ne $BackupFolderName -and $_.Extension -ne '.log'
        }
        
        if ($itemsToBackup.Count -eq 0) {
            Write-Log "No files to backup" -Level INFO
            return $true
        }
        
        Write-Log "Found $($itemsToBackup.Count) items to backup" -Level INFO
        
        $movedCount = 0
        foreach ($item in $itemsToBackup) {
            try {
                $destPath = Join-Path $BackupPath $item.Name
                Move-Item -Path $item.FullName -Destination $destPath -Force -ErrorAction Stop
                $movedCount++
                Write-Verbose "Moved: $($item.Name)"
            }
            catch {
                Write-Log "Failed to move $($item.Name): $($_.Exception.Message)" -Level WARNING
            }
        }
        
        Write-Log "Moved $movedCount items to backup folder" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Error moving files to backup: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ============================================================================
# CLEANUP FUNCTIONS
# ============================================================================
function Remove-OldTPSFiles {
    param(
        [string]$TPSFolder,
        [string]$BackupFolderName
    )
    
    Write-Log "Removing old TPS files (excluding backup folder and log files)..." -Level INFO
    
    try {
        # Get all items except backup folder and log files
        $itemsToRemove = Get-ChildItem -Path $TPSFolder -Force | Where-Object {
            $_.Name -ne $BackupFolderName -and $_.Extension -ne '.log'
        }
        
        if ($itemsToRemove.Count -eq 0) {
            Write-Log "No files to remove" -Level INFO
            return $true
        }
        
        Write-Log "Found $($itemsToRemove.Count) items to remove" -Level INFO
        
        $removedCount = 0
        foreach ($item in $itemsToRemove) {
            try {
                Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
                $removedCount++
                Write-Verbose "Removed: $($item.Name)"
            }
            catch {
                Write-Log "Failed to remove $($item.Name): $($_.Exception.Message)" -Level WARNING
            }
        }
        
        Write-Log "Removed $removedCount items from TPS folder" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Error removing old files: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ============================================================================
# COPY FUNCTIONS
# ============================================================================
function Copy-NewTPSFiles {
    param(
        [string]$Source,
        [string]$Destination
    )
    
    Write-Log "Copying new TPS files from source..." -Level INFO
    Write-Log "Source: $Source" -Level INFO
    Write-Log "Destination: $Destination" -Level INFO
    
    try {
        # Get all files from source
        $sourceFiles = Get-ChildItem -Path $Source -Recurse -File
        
        if ($sourceFiles.Count -eq 0) {
            Write-Log "No files found in source location" -Level WARNING
            return $false
        }
        
        $totalSize = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
        Write-Log "Files to copy: $($sourceFiles.Count) (Total size: $([math]::Round($totalSize/1MB, 2)) MB)" -Level INFO
        
        $copiedCount = 0
        $current = 0
        
        foreach ($file in $sourceFiles) {
            $current++
            $relativePath = $file.FullName.Substring($Source.Length).TrimStart('\')
            $destFile = Join-Path $Destination $relativePath
            $destDir = Split-Path $destFile -Parent
            
            # Show progress
            $percentComplete = [math]::Round(($current / $sourceFiles.Count) * 100, 2)
            Write-Progress -Activity "Copying TPS Files" `
                          -Status "Processing: $($file.Name) ($current of $($sourceFiles.Count))" `
                          -PercentComplete $percentComplete
            
            try {
                # Ensure destination directory exists
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force -ErrorAction Stop
                $copiedCount++
                Write-Verbose "Copied: $relativePath"
            }
            catch {
                Write-Log "Failed to copy $($file.Name): $($_.Exception.Message)" -Level WARNING
            }
        }
        
        Write-Progress -Activity "Copying TPS Files" -Completed
        
        Write-Log "Copied $copiedCount of $($sourceFiles.Count) files" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Error copying files: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ============================================================================
# MAIN DEPLOYMENT FUNCTION
# ============================================================================
function Start-TPSDeployment {
    try {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║           TPS DEPLOYMENT SCRIPT                            ║" -ForegroundColor Cyan
        Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
        Write-Host "║  Source: $($SourcePath.PadRight(48)) ║" -ForegroundColor Cyan
        Write-Host "║  Destination: $($DestinationPath.PadRight(43)) ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        
        # Step 1: Create destination folder if it doesn't exist
        if (-not (Test-Path $DestinationPath)) {
            Write-Log "Creating destination folder: $DestinationPath" -Level INFO
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Step 2: Initialize logging
        if (-not (Initialize-LogFile -LogPath $DestinationPath)) {
            throw "Failed to initialize logging"
        }
        
        Write-Log "Source Path: $SourcePath" -Level INFO
        Write-Log "Destination Path: $DestinationPath" -Level INFO
        Write-Log "Backup Folder: $BackupFolderName" -Level INFO
        Write-Log "" -Level INFO
        
        # Step 3: Verify source exists
        Write-Log "Step 1: Verifying source path..." -Level INFO
        if (-not (Test-Path $SourcePath)) {
            throw "Source path not found: $SourcePath"
        }
        Write-Log "Source path verified" -Level SUCCESS
        Write-Log "" -Level INFO
        
        # Step 4: Setup backup folder path
        $script:BackupFolder = Join-Path $DestinationPath $BackupFolderName
        
        # Step 5: Clean backup folder if it contains files
        Write-Log "Step 2: Preparing backup folder..." -Level INFO
        Clear-BackupFolder -BackupPath $script:BackupFolder
        Write-Log "" -Level INFO
        
        # Step 6: Move current TPS files to backup
        Write-Log "Step 3: Moving current files to backup..." -Level INFO
        if (-not (Move-FilesToBackup -SourceFolder $DestinationPath -BackupPath $script:BackupFolder)) {
            throw "Failed to move files to backup"
        }
        Write-Log "" -Level INFO
        
        # Step 7: Remove old files (this should be minimal since we moved them)
        Write-Log "Step 4: Cleaning remaining old files..." -Level INFO
        Remove-OldTPSFiles -TPSFolder $DestinationPath -BackupFolderName $BackupFolderName
        Write-Log "" -Level INFO
        
        # Step 8: Copy new files from source
        Write-Log "Step 5: Copying new files from source..." -Level INFO
        if (-not (Copy-NewTPSFiles -Source $SourcePath -Destination $DestinationPath)) {
            throw "Failed to copy new files"
        }
        Write-Log "" -Level INFO
        
        # Step 9: Deployment summary
        $duration = (Get-Date) - $script:StartTime
        
        Write-Log "========================================" -Level SUCCESS
        Write-Log "DEPLOYMENT COMPLETED SUCCESSFULLY" -Level SUCCESS
        Write-Log "========================================" -Level SUCCESS
        Write-Log "Duration: $($duration.ToString('hh\:mm\:ss'))" -Level INFO
        Write-Log "Log file: $script:LogFile" -Level INFO
        Write-Log "" -Level INFO
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║           DEPLOYMENT COMPLETED SUCCESSFULLY                ║" -ForegroundColor Green
        Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Green
        Write-Host "║  Duration: $($duration.ToString('hh\:mm\:ss').PadRight(48)) ║" -ForegroundColor Green
        Write-Host "║  Log File: $([System.IO.Path]::GetFileName($script:LogFile).PadRight(48)) ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Log "========================================" -Level ERROR
        Write-Log "DEPLOYMENT FAILED" -Level ERROR
        Write-Log "Error: $($_.Exception.Message)" -Level ERROR
        Write-Log "========================================" -Level ERROR
        
        Write-Host ""
        Write-Host "╔═════════════════════════════════════════════════════════��══╗" -ForegroundColor Red
        Write-Host "║           DEPLOYMENT FAILED                                ║" -ForegroundColor Red
        Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Red
        Write-Host "║  Check log file for details                                ║" -ForegroundColor Red
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        
        throw
    }
}

# ============================================================================
# SCRIPT EXECUTION
# ============================================================================
try {
    Start-TPSDeployment
}
catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    exit 1
}
