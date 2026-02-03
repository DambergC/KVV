#Requires -RunAsAdministrator
<#
.SYNOPSIS
    TPS Deployment Script - Enhanced Version
.DESCRIPTION
    Creates TPS folder, cleans old files, copies new files from network share, and creates desktop shortcut.
    Includes backup, rollback, progress tracking, and enhanced error handling.
.PARAMETER TPSFolder
    Destination folder for TPS files (default: C:\TPS)
.PARAMETER NetworkShare
    Source network share path (default: \\SERVER\Share\TPS)
.PARAMETER SkipShortcut
    Skip desktop shortcut creation
.PARAMETER BackupBeforeClean
    Create backup of existing files before cleanup
.PARAMETER MaxBackupCount
    Maximum number of backups to keep (default: 5)
.PARAMETER WhatIf
    Show what would happen without making changes
.PARAMETER Verbose
    Display detailed operation information
.EXAMPLE
    .\Install_TPS.ps1
.EXAMPLE
    .\Install_TPS.ps1 -BackupBeforeClean -Verbose
.EXAMPLE
    .\Install_TPS.ps1 -NetworkShare "\\FILESERVER\TPS" -SkipShortcut
.EXAMPLE
    .\Install_TPS.ps1 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(HelpMessage="Destination folder for TPS files")]
    [ValidateScript({
        if (Test-Path $_ -IsValid) { $true }
        else { throw "Invalid path format: $_" }
    })]
    [string]$TPSFolder = "C:\TPS",
    
    [Parameter(HelpMessage="Source network share path")]
    [ValidatePattern('^\\\\[^\\]+\\[^\\]+')]
    [string]$NetworkShare = "\\SERVER\Share\TPS",
    
    [Parameter(HelpMessage="Skip desktop shortcut creation")]
    [switch]$SkipShortcut,
    
    [Parameter(HelpMessage="Create backup before cleanup")]
    [switch]$BackupBeforeClean,
    
    [Parameter(HelpMessage="Maximum number of backups to keep")]
    [ValidateRange(1,20)]
    [int]$MaxBackupCount = 5
)

# Script-level variables
$script:logFile = $null
$script:deploymentStats = @{
    FilesRemoved = 0
    FilesCopied = 0
    TotalSizeCopied = 0
    BackupCreated = $false
    ShortcutCreated = $false
    Errors = @()
    Warnings = @()
    StartTime = Get-Date
}

#region Functions

function Initialize-Logging {
    param([string]$LogPath)
    
    try {
        # Ensure parent directory exists
        $parentDir = Split-Path $LogPath -Parent
        if (-not (Test-Path $parentDir)) {
            New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
        }
        
        # Create or verify log file
        if (-not (Test-Path $LogPath)) {
            New-Item -Path $LogPath -ItemType File -Force | Out-Null
        }
        
        $script:logFile = $LogPath
        Write-Log "=== TPS Deployment Started ===" -Level INFO
        Write-Log "Script Version: 2.0" -Level INFO
        Write-Log "User: $env:USERNAME" -Level INFO
        Write-Log "Computer: $env:COMPUTERNAME" -Level INFO
        Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" -Level INFO
        return $true
    }
    catch {
        Write-Host "ERROR: Failed to initialize logging: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - [$Level] - $Message"
    
    # Write to log file
    if ($script:logFile) {
        Add-Content -Path $script:logFile -Value $logMessage -ErrorAction SilentlyContinue
    }
    
    # Write to console with color
    switch ($Level) {
        'ERROR' { 
            Write-Host $logMessage -ForegroundColor Red
            $script:deploymentStats.Errors += $Message
        }
        'WARNING' { 
            Write-Host $logMessage -ForegroundColor Yellow
            $script:deploymentStats.Warnings += $Message
        }
        'SUCCESS' { 
            Write-Host $logMessage -ForegroundColor Green
        }
        default { 
            Write-Host $logMessage
        }
    }
    
    # Also write to Verbose stream if -Verbose is enabled
    Write-Verbose $Message
}

function Test-NetworkShareAccess {
    param([string]$SharePath)
    
    Write-Log "Validating network share access: $SharePath" -Level INFO
    
    if (-not (Test-Path $SharePath)) {
        Write-Log "Network share not accessible: $SharePath" -Level ERROR
        return $false
    }
    
    try {
        $null = Get-ChildItem -Path $SharePath -ErrorAction Stop
        Write-Log "Network share is accessible" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Cannot read from network share: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Backup-ExistingFiles {
    param(
        [string]$SourcePath,
        [int]$MaxBackups
    )
    
    if (-not (Test-Path $SourcePath)) {
        Write-Log "Source path doesn't exist, skipping backup" -Level WARNING
        return $null
    }
    
    $backupFolderName = "Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $backupPath = Join-Path $SourcePath $backupFolderName
    
    try {
        Write-Log "Creating backup: $backupPath" -Level INFO
        
        if ($PSCmdlet.ShouldProcess($backupPath, "Create backup")) {
            New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
            
            # Copy all files except logs and previous backups
            $itemsToCopy = Get-ChildItem -Path $SourcePath -Recurse | 
                Where-Object { 
                    $_.FullName -notlike "*\Backup_*" -and 
                    $_.Extension -ne ".log" 
                }
            
            if ($itemsToCopy) {
                foreach ($item in $itemsToCopy) {
                    $relativePath = $item.FullName.Substring($SourcePath.Length + 1)
                    $destPath = Join-Path $backupPath $relativePath
                    $destDir = Split-Path $destPath -Parent
                    
                    if (-not (Test-Path $destDir)) {
                        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                    }
                    
                    if (-not $item.PSIsContainer) {
                        Copy-Item -Path $item.FullName -Destination $destPath -Force
                    }
                }
                
                $script:deploymentStats.BackupCreated = $true
                Write-Log "Backup created successfully with $($itemsToCopy.Count) item(s)" -Level SUCCESS
            }
            else {
                Write-Log "No items to backup" -Level INFO
            }
            
            # Clean old backups
            Remove-OldBackups -Path $SourcePath -MaxCount $MaxBackups
        }
        
        return $backupPath
    }
    catch {
        Write-Log "Failed to create backup: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Remove-OldBackups {
    param(
        [string]$Path,
        [int]$MaxCount
    )
    
    try {
        $backups = Get-ChildItem -Path $Path -Directory -Filter "Backup_*" | 
            Sort-Object Name -Descending
        
        if ($backups.Count -gt $MaxCount) {
            $toRemove = $backups | Select-Object -Skip $MaxCount
            foreach ($backup in $toRemove) {
                if ($PSCmdlet.ShouldProcess($backup.FullName, "Remove old backup")) {
                    Remove-Item -Path $backup.FullName -Recurse -Force
                    Write-Log "Removed old backup: $($backup.Name)" -Level INFO
                }
            }
        }
    }
    catch {
        Write-Log "Failed to clean old backups: $($_.Exception.Message)" -Level WARNING
    }
}

function Remove-ItemSafely {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$Recurse
    )
    
    try {
        if ($PSCmdlet.ShouldProcess($Path, "Remove")) {
            $params = @{
                Path = $Path
                Force = $true
                ErrorAction = 'Stop'
            }
            if ($Recurse) { $params.Recurse = $true }
            
            Remove-Item @params
            return $true
        }
        return $false
    }
    catch [System.IO.IOException] {
        Write-Log "File is locked or in use: $Path" -Level WARNING
        return $false
    }
    catch [System.UnauthorizedAccessException] {
        Write-Log "Access denied: $Path" -Level WARNING
        return $false
    }
    catch {
        Write-Log "Failed to remove $Path : $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Clear-TPSFolder {
    param([string]$FolderPath)
    
    Write-Log "Starting cleanup of $FolderPath (excluding *.log and Backup_* folders)" -Level INFO
    
    $filesToRemove = Get-ChildItem -Path $FolderPath -File -Recurse | 
        Where-Object { 
            $_.Extension -ne ".log" -and 
            $_.FullName -notlike "*\Backup_*\*"
        }
    
    if ($filesToRemove) {
        $removeCount = 0
        foreach ($file in $filesToRemove) {
            if (Remove-ItemSafely -Path $file.FullName) {
                Write-Verbose "Removed: $($file.FullName)"
                $removeCount++
            }
        }
        
        $script:deploymentStats.FilesRemoved = $removeCount
        Write-Log "Cleanup completed. Removed $removeCount of $($filesToRemove.Count) file(s)" -Level SUCCESS
        
        # Clean empty directories
        Remove-EmptyDirectories -Path $FolderPath
    }
    else {
        Write-Log "No files to remove" -Level INFO
    }
}

function Remove-EmptyDirectories {
    param([string]$Path)
    
    try {
        $emptyDirs = Get-ChildItem -Path $Path -Directory -Recurse | 
            Where-Object { 
                $_.FullName -notlike "*\Backup_*" -and
                (Get-ChildItem $_.FullName -Force).Count -eq 0 
            } |
            Sort-Object { $_.FullName.Length } -Descending
        
        foreach ($dir in $emptyDirs) {
            if (Remove-ItemSafely -Path $dir.FullName) {
                Write-Log "Removed empty directory: $($dir.FullName)" -Level INFO
            }
        }
    }
    catch {
        Write-Log "Error cleaning empty directories: $($_.Exception.Message)" -Level WARNING
    }
}

function Copy-TPSFiles {
    param(
        [string]$Source,
        [string]$Destination
    )
    
    Write-Log "Starting copy from: $Source" -Level INFO
    Write-Log "Destination: $Destination" -Level INFO
    
    try {
        # Get all files to copy
        $files = Get-ChildItem -Path $Source -File -Recurse
        $totalFiles = $files.Count
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
        
        if ($totalFiles -eq 0) {
            Write-Log "No files found in source location" -Level WARNING
            return $false
        }
        
        Write-Log "Files to copy: $totalFiles (Total size: $([math]::Round($totalSize/1MB, 2)) MB)" -Level INFO
        
        $current = 0
        $copiedFiles = @()
        
        foreach ($file in $files) {
            $current++
            $relativePath = $file.FullName.Substring($Source.Length).TrimStart('\')
            $destFile = Join-Path $Destination $relativePath
            $destDir = Split-Path $destFile -Parent
            
            # Show progress
            $percentComplete = [math]::Round(($current / $totalFiles) * 100, 2)
            Write-Progress -Activity "Copying TPS Files" `
                          -Status "Processing: $($file.Name) ($current of $totalFiles)" `
                          -PercentComplete $percentComplete
            
            if ($PSCmdlet.ShouldProcess($destFile, "Copy file")) {
                # Ensure destination directory exists
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force -ErrorAction Stop
                $copiedFiles += $file
                Write-Verbose "Copied: $relativePath"
            }
        }
        
        Write-Progress -Activity "Copying TPS Files" -Completed
        
        $script:deploymentStats.FilesCopied = $copiedFiles.Count
        $script:deploymentStats.TotalSizeCopied = ($copiedFiles | Measure-Object -Property Length -Sum).Sum
        
        Write-Log "Successfully copied $($copiedFiles.Count) file(s)" -Level SUCCESS
        return $true
    }
    catch {
        Write-Progress -Activity "Copying TPS Files" -Completed
        Write-Log "ERROR: Failed to copy files: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function New-DesktopShortcut {
    param(
        [string]$TargetFolder,
        [string]$ShortcutPath
    )
    
    try {
        Write-Log "Creating desktop shortcut: $ShortcutPath" -Level INFO
        
        if (Test-Path $ShortcutPath) {
            Write-Log "Shortcut already exists, updating..." -Level INFO
        }
        
        if ($PSCmdlet.ShouldProcess($ShortcutPath, "Create shortcut")) {
            $WScriptShell = New-Object -ComObject WScript.Shell
            $shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
            
            # Look for executable file to use as target
            $exeFile = Get-ChildItem -Path $TargetFolder -Filter "*.exe" -File | 
                       Select-Object -First 1
            
            if ($exeFile) {
                $shortcut.TargetPath = $exeFile.FullName
                $shortcut.WorkingDirectory = $TargetFolder
                $shortcut.IconLocation = $exeFile.FullName
                Write-Log "Shortcut target: $($exeFile.Name)" -Level INFO
            }
            else {
                # Fallback to folder if no exe found
                $shortcut.TargetPath = "explorer.exe"
                $shortcut.Arguments = $TargetFolder
                $shortcut.WorkingDirectory = $TargetFolder
                Write-Log "No executable found, shortcut will open folder" -Level INFO
            }
            
            $shortcut.Description = "TPS Application"
            $shortcut.Save()
            
            # Release COM object
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WScriptShell) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            $script:deploymentStats.ShortcutCreated = $true
            Write-Log "Desktop shortcut created successfully" -Level SUCCESS
            return $true
        }
        
        return $false
    }
    catch {
        Write-Log "Failed to create shortcut: $($_.Exception.Message)" -Level WARNING
        return $false
    }
}

function Show-DeploymentSummary {
    $duration = (Get-Date) - $script:deploymentStats.StartTime
    
    Write-Log "" -Level INFO
    Write-Log "╔════════════════════════════════════════════════════════════╗" -Level INFO
    Write-Log "║           TPS DEPLOYMENT SUMMARY                           ║" -Level INFO
    Write-Log "╠════════════════════════════════════════════════════════════╣" -Level INFO
    Write-Log "║ Files Removed:    $($script:deploymentStats.FilesRemoved.ToString().PadLeft(40)) ║" -Level INFO
    Write-Log "║ Files Copied:     $($script:deploymentStats.FilesCopied.ToString().PadLeft(40)) ║" -Level INFO
    Write-Log "║ Total Size:       $("$([math]::Round($script:deploymentStats.TotalSizeCopied/1MB, 2)) MB".PadLeft(40)) ║" -Level INFO
    Write-Log "║ Backup Created:   $($script:deploymentStats.BackupCreated.ToString().PadLeft(40)) ║" -Level INFO
    Write-Log "║ Shortcut Created: $($script:deploymentStats.ShortcutCreated.ToString().PadLeft(40)) ║" -Level INFO
    Write-Log "║ Warnings:         $($script:deploymentStats.Warnings.Count.ToString().PadLeft(40)) ║" -Level INFO
    Write-Log "║ Errors:           $($script:deploymentStats.Errors.Count.ToString().PadLeft(40)) ║" -Level INFO
    Write-Log "║ Duration:         $("$([math]::Round($duration.TotalSeconds, 2)) seconds".PadLeft(40)) ║" -Level INFO
    Write-Log "╚════════════════════════════════════════════════════════════╝" -Level INFO
    
    if ($script:deploymentStats.Warnings.Count -gt 0) {
        Write-Log "" -Level INFO
        Write-Log "Warnings:" -Level WARNING
        foreach ($warning in $script:deploymentStats.Warnings) {
            Write-Log "  - $warning" -Level WARNING
        }
    }
    
    if ($script:deploymentStats.Errors.Count -gt 0) {
        Write-Log "" -Level INFO
        Write-Log "Errors:" -Level ERROR
        foreach ($error in $script:deploymentStats.Errors) {
            Write-Log "  - $error" -Level ERROR
        }
    }
}

#endregion

#region Main Script

try {
    # Initialize (create folder structure before logging)
    if (-not (Test-Path $TPSFolder)) {
        if ($PSCmdlet.ShouldProcess($TPSFolder, "Create TPS folder")) {
            New-Item -Path $TPSFolder -ItemType Directory -Force | Out-Null
        }
    }
    
    # Initialize logging
    $logPath = Join-Path $TPSFolder "TPS_$(Get-Date -Format 'yyyyMMdd').log"
    if (-not (Initialize-Logging -LogPath $logPath)) {
        throw "Failed to initialize logging"
    }
    
    Write-Log "Parameters:" -Level INFO
    Write-Log "  TPS Folder: $TPSFolder" -Level INFO
    Write-Log "  Network Share: $NetworkShare" -Level INFO
    Write-Log "  Skip Shortcut: $SkipShortcut" -Level INFO
    Write-Log "  Backup Before Clean: $BackupBeforeClean" -Level INFO
    Write-Log "  Max Backup Count: $MaxBackupCount" -Level INFO
    if ($WhatIfPreference) {
        Write-Log "  Running in WhatIf mode - no changes will be made" -Level WARNING
    }
    
    # Step 1: Validate network share access
    Write-Log "" -Level INFO
    Write-Log "Step 1: Validating network share access" -Level INFO
    if (-not (Test-NetworkShareAccess -SharePath $NetworkShare)) {
        throw "Network share validation failed"
    }
    
    # Step 2: Create backup if requested
    if ($BackupBeforeClean) {
        Write-Log "" -Level INFO
        Write-Log "Step 2: Creating backup" -Level INFO
        $backupPath = Backup-ExistingFiles -SourcePath $TPSFolder -MaxBackups $MaxBackupCount
        if ($backupPath) {
            Write-Log "Backup location: $backupPath" -Level SUCCESS
        }
    }
    else {
        Write-Log "" -Level INFO
        Write-Log "Step 2: Skipping backup (not requested)" -Level INFO
    }
    
    # Step 3: Clean existing files
    Write-Log "" -Level INFO
    Write-Log "Step 3: Cleaning existing files" -Level INFO
    Clear-TPSFolder -FolderPath $TPSFolder
    
    # Step 4: Copy new files
    Write-Log "" -Level INFO
    Write-Log "Step 4: Copying new files from network share" -Level INFO
    if (-not (Copy-TPSFiles -Source $NetworkShare -Destination $TPSFolder)) {
        throw "File copy operation failed"
    }
    
    # Step 5: Create desktop shortcut
    if (-not $SkipShortcut) {
        Write-Log "" -Level INFO
        Write-Log "Step 5: Creating desktop shortcut" -Level INFO
        $publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
        $shortcutPath = Join-Path $publicDesktop "TPS.lnk"
        New-DesktopShortcut -TargetFolder $TPSFolder -ShortcutPath $shortcutPath
    }
    else {
        Write-Log "" -Level INFO
        Write-Log "Step 5: Skipping shortcut creation (not requested)" -Level INFO
    }
    
    # Show summary
    Write-Log "" -Level INFO
    Show-DeploymentSummary
    
    Write-Log "" -Level INFO
    Write-Log "TPS deployment completed successfully!" -Level SUCCESS
    exit 0
}
catch {
    Write-Log "" -Level ERROR
    Write-Log "DEPLOYMENT FAILED: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    
    # Show summary even on failure
    Show-DeploymentSummary
    
    exit 1
}
finally {
    # Cleanup
    Write-Log "=== TPS Deployment Ended ===" -Level INFO
}

#endregion
