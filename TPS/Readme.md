# TPS Deployment Script

Automated PowerShell script for deploying TPS application files from a network share to local machines with backup, rollback capabilities, and comprehensive logging.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Parameters](#parameters)
- [Usage Examples](#usage-examples)
- [Scheduling with Task Scheduler](#scheduling-with-task-scheduler)
- [File Structure](#file-structure)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)
- [Best Practices](#best-practices)

## 🎯 Overview

The TPS Deployment Script (`Run-TPS.ps1`) automates the process of:
1. Creating and managing the local TPS folder
2. Backing up existing files (optional)
3. Cleaning old files while preserving logs
4. Copying new files from a network share
5. Creating desktop shortcuts for easy access
6. Providing detailed logging and deployment statistics

## ✨ Features

- **✅ Backup & Rollback**: Automatic backup before deployment with configurable retention
- **✅ Progress Tracking**: Real-time progress bars during file operations
- **✅ Enhanced Error Handling**: Graceful handling of locked files and access issues
- **✅ Comprehensive Logging**: Timestamped logs with severity levels
- **✅ WhatIf Support**: Preview changes before executing
- **✅ Network Share Validation**: Verifies access before proceeding
- **✅ Desktop Shortcut Creation**: Automatic shortcut with smart executable detection
- **✅ Deployment Statistics**: Detailed summary of operations performed

## 📦 Prerequisites

### System Requirements
- **Operating System**: Windows 10/11, Windows Server 2016+
- **PowerShell**: Version 5.1 or higher
- **Permissions**: Administrator rights (script enforces `-RunAsAdministrator`)
- **Network Access**: Access to the source network share

### Required Permissions
- Local Administrator rights
- Read access to network share
- Write access to destination folder (default: `C:\TPS`)

## ⚙️ Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TPSFolder` | String | `C:\TPS` | Destination folder for TPS files |
| `NetworkShare` | String | `\\SERVER\Share\TPS` | Source network share path |
| `SkipShortcut` | Switch | `$false` | Skip desktop shortcut creation |
| `BackupBeforeClean` | Switch | `$false` | Create backup before cleanup |
| `MaxBackupCount` | Int | `5` | Maximum number of backups to keep (1-20) |
| `WhatIf` | Switch | `$false` | Preview changes without executing |
| `Verbose` | Switch | `$false` | Display detailed operation information |

## 🚀 Usage Examples

### Basic Usage
```powershell
# Run with default settings
.\Run-TPS.ps1
```

### With Backup
```powershell
# Create backup before deployment
.\Run-TPS.ps1 -BackupBeforeClean
```

### Custom Network Share
```powershell
# Specify custom network share
.\Run-TPS.ps1 -NetworkShare "\\FILESERVER\Applications\TPS" -BackupBeforeClean
```

### Preview Mode (WhatIf)
```powershell
# See what would happen without making changes
.\Run-TPS.ps1 -WhatIf
```

### Verbose Output
```powershell
# Detailed operation information
.\Run-TPS.ps1 -BackupBeforeClean -Verbose
```

### Skip Shortcut Creation
```powershell
# Deploy without creating desktop shortcut
.\Run-TPS.ps1 -SkipShortcut -BackupBeforeClean
```

### Custom Destination and Backup Settings
```powershell
# Full customization
.\Run-TPS.ps1 -TPSFolder "D:\Applications\TPS" `
              -NetworkShare "\\SERVER\TPS" `
              -BackupBeforeClean `
              -MaxBackupCount 10 `
              -Verbose
```

## ⏰ Scheduling with Task Scheduler

### Method 1: Task Scheduler GUI

#### Step-by-Step Instructions

1. **Open Task Scheduler**
   - Press `Win + R`
   - Type `taskschd.msc`
   - Press Enter

2. **Create New Task**
   - Click **"Create Task"** (not "Create Basic Task")

3. **General Tab Configuration**
   ```
   Name: TPS Deployment - Hourly
   Description: Runs TPS deployment script every hour
   ☑ Run whether user is logged on or not
   ☑ Run with highest privileges
   Configure for: Windows 10 / Windows Server 2016
   ```

4. **Triggers Tab Configuration**
   - Click **"New..."**
   - Begin the task: **On a schedule**
   - Settings: **Daily**
   - Recur every: **1 days**
   - ☑ **Repeat task every: 1 hour**
   - For a duration of: **Indefinitely**
   - ☑ **Enabled**

5. **Actions Tab Configuration**
   - Click **"New..."**
   - Action: **Start a program**
   - Program/script:
     ```
     powershell.exe
     ```
   - Add arguments:
     ```
     -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "C:\TPS\Run-TPS.ps1" -BackupBeforeClean
     ```
   - Start in:
     ```
     C:\TPS
     ```

6. **Conditions Tab** (optional)
   ```
   ☐ Start only if on AC power (uncheck for laptops)
   ☑ Wake the computer to run this task
   ☑ Start only if the following network connection is available: Any connection
   ```

7. **Settings Tab**
   ```
   ☑ Allow task to be run on demand
   ☑ Run task as soon as possible after a scheduled start is missed
   ☑ If the task fails, restart every: 10 minutes (3 attempts)
   ☐ Stop the task if it runs longer than: (leave unchecked)
   If the running task does not end when requested: Stop the existing instance
   ```

8. **Save and Test**
   - Click **OK**
   - Enter administrator credentials
   - Right-click the task → **Run** to test

---

### Method 2: PowerShell Script

Create the scheduled task automatically using PowerShell:

```powershell
# ===================================================================
# Create-TPSScheduledTask.ps1
# Run this script as Administrator to create the scheduled task
# ===================================================================

# Configuration
$taskName = "TPS Deployment - Hourly"
$scriptPath = "C:\TPS\Run-TPS.ps1"  # Update this path
$taskDescription = "Runs TPS deployment script every hour with backup"

# Verify script exists
if (-not (Test-Path $scriptPath)) {
    Write-Error "Script not found: $scriptPath"
    exit 1
}

# Create the action
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$scriptPath`" -BackupBeforeClean -MaxBackupCount 5" `
    -WorkingDirectory (Split-Path $scriptPath -Parent)

# Create the trigger (every hour, starting now)
$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Hours 1) `
    -RepetitionDuration ([TimeSpan]::MaxValue)

# Create the principal (run as SYSTEM with highest privileges)
$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

# Create settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10)

# Register the task
try {
    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Description $taskDescription `
        -Force

    Write-Host "✅ Scheduled task created successfully!" -ForegroundColor Green
    Write-Host "Task Name: $taskName" -ForegroundColor Cyan
    Write-Host "Schedule: Every hour" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To test the task, run:" -ForegroundColor Yellow
    Write-Host "  Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor White
}
catch {
    Write-Error "Failed to create scheduled task: $_"
    exit 1
}
```

**To use this script:**
```powershell
# Save as Create-TPSScheduledTask.ps1 and run as Administrator
.\Create-TPSScheduledTask.ps1
```

---

### Method 3: Business Hours Only

Schedule the task to run every hour during business hours (8 AM - 6 PM):

```powershell
# Configuration
$taskName = "TPS Deployment - Business Hours"
$scriptPath = "C:\TPS\Run-TPS.ps1"

# Create action
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$scriptPath`" -BackupBeforeClean" `
    -WorkingDirectory (Split-Path $scriptPath -Parent)

# Create multiple triggers for each hour (8 AM to 6 PM)
$triggers = @()
for ($hour = 8; $hour -le 18; $hour++) {
    $triggers += New-ScheduledTaskTrigger -Daily -At "$($hour):00"
}

# Principal and settings
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Register the task
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $triggers `
    -Principal $principal `
    -Settings $settings `
    -Description "Runs TPS deployment every hour from 8 AM to 6 PM"

Write-Host "✅ Business hours task created successfully!" -ForegroundColor Green
```

---

### Method 4: Using a Service Account

For better network share access, use a domain service account:

```powershell
# Using a service account with domain credentials
$taskName = "TPS Deployment - Hourly"
$scriptPath = "C:\TPS\Run-TPS.ps1"
$serviceAccount = "DOMAIN\svc_tps_deploy"
$servicePassword = Read-Host "Enter password for $serviceAccount" -AsSecureString

# Convert SecureString to plain text for Register-ScheduledTask
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($servicePassword)
$plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Create action
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$scriptPath`" -BackupBeforeClean" `
    -WorkingDirectory (Split-Path $scriptPath -Parent)

# Create trigger
$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Hours 1) `
    -RepetitionDuration ([TimeSpan]::MaxValue)

# Create principal with service account
$principal = New-ScheduledTaskPrincipal `
    -UserId $serviceAccount `
    -LogonType Password `
    -RunLevel Highest

# Create settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

# Register with credentials
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -User $serviceAccount `
    -Password $plainPassword `
    -Force

# Clear password from memory
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
Remove-Variable plainPassword

Write-Host "✅ Task created with service account!" -ForegroundColor Green
```

---

### Managing Scheduled Tasks

#### View Task Status
```powershell
# Get task information
Get-ScheduledTask -TaskName "TPS Deployment - Hourly" | Get-ScheduledTaskInfo

# View last run result
Get-ScheduledTask -TaskName "TPS Deployment - Hourly" | Select-Object -ExpandProperty LastRunTime, LastTaskResult
```

#### Test Task Manually
```powershell
# Run the task immediately
Start-ScheduledTask -TaskName "TPS Deployment - Hourly"
```

#### Disable/Enable Task
```powershell
# Disable task
Disable-ScheduledTask -TaskName "TPS Deployment - Hourly"

# Enable task
Enable-ScheduledTask -TaskName "TPS Deployment - Hourly"
```

#### Remove Task
```powershell
# Remove the scheduled task
Unregister-ScheduledTask -TaskName "TPS Deployment - Hourly" -Confirm:$false
```

#### Export/Import Task
```powershell
# Export task to XML
Export-ScheduledTask -TaskName "TPS Deployment - Hourly" | Out-File "C:\Backup\TPSTask.xml"

# Import task from XML
Register-ScheduledTask -Xml (Get-Content "C:\Backup\TPSTask.xml" | Out-String) -TaskName "TPS Deployment - Hourly"
```

---

## 📁 File Structure

After deployment, your TPS folder structure will look like this:

```
C:\TPS\
│
├── Run-TPS.ps1                     # Main deployment script
├── TPS_20260123.log                # Daily log file
├── TPS_20260122.log                # Previous day log
│
├── Backup_20260123_143022\         # Automatic backup (if enabled)
│   └── [previous TPS files]
│
├── Backup_20260123_120000\         # Older backup
│   └── [previous TPS files]
│
└── [TPS Application Files]         # Deployed application files
    ├── TPS.exe
    ├── config.xml
    └── data\
        └── ...
```

**Desktop Shortcut:**
```
C:\Users\Public\Desktop\TPS.lnk    # Public desktop shortcut
```

---

## 📊 Logging

### Log File Location
- **Path**: `C:\TPS\TPS_YYYYMMDD.log`
- **Format**: Daily logs (one file per day)
- **Retention**: Logs are preserved indefinitely (not cleaned by script)

### Log Entry Format
```
2026-01-23 14:30:15 - [INFO] - Step 1: Validating network share access
2026-01-23 14:30:16 - [SUCCESS] - Network share is accessible
2026-01-23 14:30:17 - [WARNING] - File is locked or in use: C:\TPS\old-file.dll
2026-01-23 14:30:45 - [ERROR] - Failed to copy file: Access denied
```

### Log Levels
- **INFO**: General information about operations
- **SUCCESS**: Successful operations
- **WARNING**: Non-critical issues (e.g., locked files)
- **ERROR**: Critical errors that affect deployment

### Viewing Logs

```powershell
# View today's log
Get-Content "C:\TPS\TPS_$(Get-Date -Format 'yyyyMMdd').log"

# View last 50 lines
Get-Content "C:\TPS\TPS_$(Get-Date -Format 'yyyyMMdd').log" -Tail 50

# Follow log in real-time (PowerShell 7+)
Get-Content "C:\TPS\TPS_$(Get-Date -Format 'yyyyMMdd').log" -Wait

# Search for errors
Get-Content "C:\TPS\TPS_$(Get-Date -Format 'yyyyMMdd').log" | Select-String "ERROR"

# View deployment summary
Get-Content "C:\TPS\TPS_$(Get-Date -Format 'yyyyMMdd').log" | Select-String "DEPLOYMENT SUMMARY" -Context 0,15
```

---

## 🔧 Troubleshooting

### Common Issues and Solutions

#### 1. Network Share Access Denied

**Symptoms:**
```
ERROR: Network share not accessible: \\SERVER\Share\TPS
```

**Solutions:**
- Verify network share path is correct
- Check that the account running the task has read permissions
- If using SYSTEM account, ensure computer account has access
- Consider using a service account with explicit permissions
- Test access manually:
  ```powershell
  Test-Path "\\SERVER\Share\TPS"
  Get-ChildItem "\\SERVER\Share\TPS"
  ```

#### 2. Access Denied to C:\TPS

**Symptoms:**
```
ERROR: Failed to create backup: Access to the path 'C:\TPS' is denied
```

**Solutions:**
- Ensure task is running with elevated privileges (Administrator)
- Check "Run with highest privileges" in Task Scheduler
- Verify NTFS permissions on C:\TPS folder
- Run manually as Administrator to test:
  ```powershell
  Start-Process powershell -Verb RunAs -ArgumentList "-File C:\TPS\Run-TPS.ps1"
  ```

#### 3. Locked Files Cannot Be Deleted

**Symptoms:**
```
WARNING: File is locked or in use: C:\TPS\app.exe
```

**Solutions:**
- This is expected if TPS application is running
- Schedule deployment during off-hours
- Close TPS application before deployment
- Script will continue with other files (non-critical warning)

#### 4. Task Runs But Nothing Happens

**Symptoms:**
- Task shows "Success" but no files are copied
- No log file created

**Solutions:**
- Check Task History in Task Scheduler
- Verify "Start in" directory is set correctly
- Check execution policy:
  ```powershell
  Get-ExecutionPolicy
  ```
- Review Task Scheduler history for error codes
- Add logging to verify task is actually running:
  ```powershell
  # Add to scheduled task arguments
  -File "C:\TPS\Run-TPS.ps1" -Verbose *> "C:\TPS\task-output.log"
  ```

#### 5. Script Requires Administrator Rights

**Symptoms:**
```
#Requires -RunAsAdministrator
```

**Solutions:**
- Ensure "Run with highest privileges" is checked in Task Scheduler
- When running manually, use "Run as Administrator"
- Verify user account is in local Administrators group

#### 6. Backup Folder Fills Disk Space

**Symptoms:**
- Multiple backup folders consuming disk space
- Deployment becomes slow

**Solutions:**
- Reduce `-MaxBackupCount` parameter (default is 5)
- Manually clean old backups:
  ```powershell
  Get-ChildItem "C:\TPS\Backup_*" | Sort-Object Name -Descending | Select-Object -Skip 3 | Remove-Item -Recurse -Force
  ```
- Schedule cleanup with separate task
- Consider storing backups on different drive

#### 7. Task Scheduler Shows "0x1" Error

**Symptoms:**
```
Last Run Result: 0x1 (Incorrect function)
```

**Solutions:**
- Check the log file for actual error
- Common causes:
  - Script threw an exception
  - Network share unavailable
  - Insufficient permissions
- Test script manually with same parameters
- Review Windows Event Viewer → Task Scheduler logs

---

### Diagnostic Commands

```powershell
# Check if script can run
Test-Path "C:\TPS\Run-TPS.ps1"
Get-ExecutionPolicy

# Verify network share
Test-Path "\\SERVER\Share\TPS"

# Check folder permissions
Get-Acl "C:\TPS" | Format-List

# View scheduled task details
Get-ScheduledTask -TaskName "TPS Deployment - Hourly" | Format-List *

# Check last run information
Get-ScheduledTaskInfo -TaskName "TPS Deployment - Hourly"

# View task history (Event Viewer)
Get-WinEvent -LogName "Microsoft-Windows-TaskScheduler/Operational" -MaxEvents 20 | 
    Where-Object {$_.Message -like "*TPS Deployment*"} | 
    Format-Table TimeCreated, Message -Wrap

# Test script with WhatIf
C:\TPS\Run-TPS.ps1 -WhatIf -Verbose
```

---

## 💡 Best Practices

### Security

1. **Use Dedicated Service Account**
   - Create domain service account (e.g., `svc_tps_deploy`)
   - Grant minimal required permissions
   - Use strong password and document in secure location

2. **Network Share Permissions**
   - Grant Read-only access to deployment account
   - Use specific share permissions, not Everyone
   - Consider using DFS for redundancy

3. **Local Folder Permissions**
   - Restrict write access to deployment account
   - Regular users should have read-only access
   - Log files should be protected

### Reliability

1. **Enable Backup**
   - Always use `-BackupBeforeClean` in production
   - Keep reasonable backup count (5-10)
   - Test restore procedure periodically

2. **Monitoring**
   - Review logs daily for errors/warnings
   - Set up alerts for deployment failures
   - Monitor disk space usage

3. **Testing**
   - Test in development environment first
   - Use `-WhatIf` to preview changes
   - Verify rollback procedure works

### Performance

1. **Schedule During Off-Hours**
   - Avoid peak usage times
   - Consider business hours restrictions
   - Allow time for deployment to complete

2. **Network Bandwidth**
   - Schedule around network maintenance
   - Consider file size and transfer time
   - Use local staging if network is unreliable

3. **Maintenance Windows**
   - Coordinate with application users
   - Document deployment schedule
   - Provide notification of changes

### Maintenance

1. **Regular Reviews**
   - Review logs monthly
   - Clean old backups if space constrained
   - Update documentation as needed

2. **Version Control**
   - Keep script in version control (Git)
   - Document changes in commit messages
   - Tag releases

3. **Documentation**
   - Document custom parameters used
   - Keep runbook updated
   - Train multiple team members

---

## 📋 Deployment Checklist

Use this checklist when setting up the scheduled deployment:

- [ ] Script tested manually with `-WhatIf`
- [ ] Network share accessible from target machine
- [ ] Service account created with appropriate permissions
- [ ] C:\TPS folder exists with correct permissions
- [ ] Scheduled task created and configured
- [ ] Task tested manually (Run now)
- [ ] Log file generated successfully
- [ ] Backup folder created (if enabled)
- [ ] Desktop shortcut created
- [ ] Task History enabled in Task Scheduler
- [ ] Monitoring/alerting configured
- [ ] Documentation updated
- [ ] Team trained on process
- [ ] Rollback procedure tested

---

## 📞 Support

### Getting Help

1. **Check Logs**: Review `C:\TPS\TPS_YYYYMMDD.log`
2. **View Task History**: Task Scheduler → Task History
3. **Test Manually**: Run script as Administrator with `-Verbose`
4. **Review Event Viewer**: Windows Logs → Application

### Reporting Issues

When reporting issues, include:
- Full error message from log file
- PowerShell version (`$PSVersionTable.PSVersion`)
- Operating System version
- Task Scheduler configuration
- Network environment details

---

## 📄 License

[Specify your license here]

---

## 👥 Contributors

- [Your Name/Team]
- [Additional Contributors]

---

## 📝 Changelog

### Version 2.0 (2026-01-23)
- Enhanced error handling with retry logic
- Added backup and rollback functionality
- Implemented progress tracking
- Added WhatIf support
- Comprehensive logging improvements
- Deployment statistics summary
- Old backup cleanup

### Version 1.0 (Initial Release)
- Basic deployment functionality
- Network share copy
- Desktop shortcut creation
- Simple logging

---

## 🔗 Related Documentation

- [PowerShell Execution Policies](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies)
- [Task Scheduler Overview](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page)
- [PowerShell Scheduled Tasks](https://docs.microsoft.com/en-us/powershell/module/scheduledtasks/)

---

**Last Updated**: 2026-01-23  
**Script Version**: 2.0  
**Author**: DambergC
