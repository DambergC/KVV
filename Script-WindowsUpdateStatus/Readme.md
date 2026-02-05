# Send-WindowsUpdateStatus

A PowerShell script for monitoring and reporting Windows update deployment status through Microsoft Endpoint Configuration Manager (MECM/SCCM).

## Overview

This script generates automated HTML reports showing the status of Windows update deployments and sends email notifications to administrators. It tracks patch deployment progress, identifies servers needing attention, and provides visual charts for quick status assessment.

## Features

- **Automated Reporting**: Runs on a scheduled basis to report patch status after Patch Tuesday
- **Visual HTML Reports**: Creates interactive HTML reports with charts and tables using PSWriteHTML
- **Email Notifications**: Sends detailed status emails with attached HTML reports
- **Multiple Deployment Support**: Monitors multiple deployment groups with different schedules
- **Patch Tuesday Aware**: Automatically calculates reporting dates based on Patch Tuesday schedule
- **Status Tracking**: Categorizes servers as "Success", "Error", "Unknown", or "In Progress"
- **Log File Management**: Automatic log rotation when files exceed threshold
- **Month Exclusion**: Ability to skip reporting for specific months

## Prerequisites

### Required Modules
- **ConfigurationManager**: SCCM/MECM PowerShell module
- **PatchManagementSupportTools**: For Patch Tuesday date calculation
- **PSWriteHTML**: For generating interactive HTML reports
- **Send-MailKitMessage**: For sending emails with attachments

### Permissions
- Read access to SCCM site server
- Permissions to query SCCM deployments and collections
- SMTP relay permissions for sending emails

### Environment
- Must run on or have access to SCCM site server
- PowerShell 5.1 or higher
- Network access to SMTP server

## Installation

1. Copy `Send-WindowsUpdateStatus.ps1` to your scripts directory
2. Copy `ScriptConfigPatchStatus.xml` to the same directory
3. Configure the XML file with your environment settings
4. Install required PowerShell modules:
```powershell
Install-Module PSWriteHTML
Install-Module PatchManagementSupportTools
Install-Module Send-MailKitMessage
```

## Configuration

### XML Configuration File: ScriptConfigPatchStatus.xml

The script requires an XML configuration file in the same directory.

#### Complete XML Template

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
    <Logfile>
        <Path>G:\Scripts\Logfiles\</Path>
        <Name>WindowsUpdateScript.log</Name>
        <Logfilethreshold>2000000</Logfilethreshold>
    </Logfile>
    <HTMLfilePath>G:\Scripts\OutFiles\</HTMLfilePath>
    <RunScript>
        <Job DeploymentID="16777362" Offsetdays="2" Description="Grupp100"/>
        <Job DeploymentID="16777363" Offsetdays="8" Description="Grupp200"/>
        <Job DeploymentID="16777364" Offsetdays="15" Description="Grupp300"/>
    </RunScript>
    <DisableReportMonth>
        <DisableReportMonth Number=""/>
        <DisableReportMonth Number=""/>
        <DisableReportMonth Number=""/>
    </DisableReportMonth>
    <Recipients>
        <Recipients email="christian.damberg@domain.org"/>
    </Recipients>
    <UpdateDeployed>
        <LimitDays>-25</LimitDays>
        <UpdateGroupName>Server Patch Tuesday</UpdateGroupName>
        <DaysAfterPatchToRun>1</DaysAfterPatchToRun>
    </UpdateDeployed>
    <SiteServer>siteserver</SiteServer>
    <Mailfrom>no-reply@domain.org</Mailfrom>
    <MailSMTP>smtp.domain.org</MailSMTP>
    <MailPort>25</MailPort>
    <MailCustomer>My Company</MailCustomer>
</Configuration>
```

#### XML Parameters Explained

| Section | Parameter | Description | Example |
|---------|-----------|-------------|---------|
| **Logfile** | Path | Directory for log files | `G:\Scripts\Logfiles\` |
| | Name | Log file name | `WindowsUpdateScript.log` |
| | Logfilethreshold | Max log size in bytes before rotation | `2000000` (2MB) |
| **HTMLfilePath** | - | Directory where HTML reports are saved | `G:\Scripts\OutFiles\` |
| **RunScript/Job** | DeploymentID | SCCM deployment ID to monitor | `16777362` |
| | Offsetdays | Days after Patch Tuesday to run report | `2` (runs 2 days after Patch Tuesday) |
| | Description | Group name/identifier | `Grupp100` |
| **DisableReportMonth** | Number | Month numbers to skip reporting (1-12) | `7` (skip July) |
| **Recipients** | email | Email addresses to receive reports | `admin@domain.org` |
| **UpdateDeployed** | LimitDays | How many days back to check updates | `-25` |
| | UpdateGroupName | Name of update group | `Server Patch Tuesday` |
| | DaysAfterPatchToRun | Days after patch to run | `1` |
| **SiteServer** | - | SCCM site server name | `siteserver.domain.org` |
| **Mailfrom** | - | Email sender address | `no-reply@domain.org` |
| **MailSMTP** | - | SMTP server address | `smtp.domain.org` |
| **MailPort** | - | SMTP port number | `25` or `587` |
| **MailCustomer** | - | Company/Customer name for reports | `My Company` |

## How It Works

### Workflow

1. **Initialization**
   - Loads XML configuration
   - Checks log file size and rotates if needed
   - Imports required SCCM and PatchManagement modules
   
2. **Date Calculation**
   - Determines current month's Patch Tuesday
   - Calculates reporting dates based on offset days
   - Checks if current date matches scheduled reporting date
   
3. **Month Skip Check**
   - Verifies if current month is in the disabled months list
   - Exits script if reporting is disabled for this month
   
4. **Data Collection** (for each Job/Deployment)
   - Queries SCCM for deployment status using DeploymentID
   - Retrieves status for all devices in the deployment collection
   - Categorizes devices by status: Success, Error, Unknown, In Progress
   
5. **Report Generation**
   - Creates interactive HTML report with:
     - Bar chart showing success vs. needs attention
     - Detailed table with all device statuses
     - Collection name and statistics
   - Saves HTML file with date stamp
   
6. **Email Notification**
   - Constructs email with summary information
   - Attaches generated HTML report
   - Sends to all configured recipients

### Scheduling Logic

Reports are generated based on **Patch Tuesday + Offset Days**:

- **Grupp100**: 2 days after Patch Tuesday
- **Grupp200**: 8 days after Patch Tuesday  
- **Grupp300**: 15 days after Patch Tuesday

This allows staggered patching groups to be monitored at appropriate intervals.

## Usage

### Manual Execution
```powershell
.\Send-WindowsUpdateStatus.ps1
```

### Scheduled Task
Create a scheduled task on the SCCM site server:

```powershell
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument '-ExecutionPolicy Bypass -File "G:\Scripts\Send-WindowsUpdateStatus.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At 8:00AM
$principal = New-ScheduledTaskPrincipal -UserId "DOMAIN\ServiceAccount" -RunLevel Highest
Register-ScheduledTask -TaskName "Windows Update Status Report" -Action $action -Trigger $trigger -Principal $principal
```

## Output

### HTML Report Contains:
- **Header**: Company name and title
- **Chart Section**: Visual bar chart showing:
  - Red bar: Servers needing attention (Error/Unknown/In Progress)
  - Green bar: Successful servers
- **Data Table**: Detailed information including:
  - Server name
  - Collection name
  - Status
  - Status timestamp
- **Footer**: Report generation date and computer name

### Email Contains:
- Company logo (embedded)
- Collection name
- Count of servers needing attention
- Count of successful servers
- HTML table showing servers that need attention
- Attached full HTML report

### Log File:
- Timestamped entries for all script actions
- Located in configured log path
- Automatically rotated when size exceeds threshold

## Functions

### `Rotate-Log`
Checks log file size and rotates to archived folder if threshold exceeded

### `Write-Log`
Writes timestamped log entries to log file

### `Get-CMSiteCode`
Retrieves SCCM site code from WMI

### `Get-PatchTuesday`
Calculates Patch Tuesday date for specified month/year

### `Get-CMModule`
Loads ConfigurationManager PowerShell module

### `Get-SCCMSoftwareUpdateStatus`
Queries SCCM for deployment status (from PatchManagementSupportTools)

## Status Codes

| Status | Description |
|--------|-------------|
| **Success** | Update successfully installed |
| **Error** | Update installation failed |
| **Unknown** | Status cannot be determined |
| **InProgress** | Update installation in progress |

Servers with Error, Unknown, or InProgress status are flagged as "Needs Attention"

## Troubleshooting

### Script Doesn't Run
- Verify XML file exists in same directory as script
- Check scheduled task credentials have proper permissions
- Review log file for error messages

### No Email Received
- Verify SMTP server settings in XML
- Check recipient email addresses are correct
- Ensure Send-MailKitMessage module is installed
- Test SMTP connectivity from server

### Wrong Date/Not Running
- Script only runs on scheduled days (Patch Tuesday + Offset)
- Check current date against calculated Patch Tuesday
- Review log file for date comparison details

### No Data in Report
- Verify DeploymentID exists in SCCM
- Check that deployment has run and devices are reporting
- Ensure service account has read permissions on SCCM

### Module Not Found
Install missing modules:
```powershell
Install-Module PSWriteHTML -Force
Install-Module PatchManagementSupportTools -Force
Install-Module Send-MailKitMessage -Force
```

## File Structure

```
Script-WindowsUpdateStatus/
│
├── Send-WindowsUpdateStatus.ps1    # Main script
├── ScriptConfigPatchStatus.xml     # Configuration file
└── README.md                        # This file
```

## Author Information

- **Created with**: SAPIEN Technologies, Inc., PowerShell Studio 2024
- **Created on**: October 16, 2023
- **Updated on**: March 25, 2024
- **Created by**: Christian Damberg
- **Organization**: Telia Cygate AB

## Notes

- Script designed to run on or from SCCM site server
- Requires appropriate SCCM permissions to query deployments
- HTML reports use embedded base64 logo image
- Log files automatically moved to OLDLOG subfolder when rotated
- Multiple recipients can be added to XML configuration
- Month numbers for DisableReportMonth: 1=January, 2=February, etc.

## Version History

| Date | Version | Changes |
|------|---------|---------|
| 2023-10-16 | 1.0 | Initial creation |
| 2024-03-25 | 1.1 | Updates and refinements |

## Related Scripts

- **Send-WindowsUpdateDeployed.ps1**: Reports on newly deployed updates
- Other Windows Update automation scripts in the KVV repository

## License

This script is provided AS IS with no warranty. Test thoroughly in a non-production environment before deploying.
