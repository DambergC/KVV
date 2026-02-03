# Send-WindowsUpdateDeployed

A PowerShell automation script for reporting Windows Update deployments through Microsoft Endpoint Configuration Manager (MECM/SCCM). The script generates and emails HTML reports of deployed updates based on configurable schedules.

## Overview

This script automatically creates reports on Windows updates for one or more deployments specified in an XML configuration file and sends email notifications to designated recipients. It's designed to run on a MECM site server either manually or as a scheduled task.

### Key Features

- 📅 **Automated Scheduling** - Runs on specified days after Patch Tuesday
- 📧 **Email Reports** - Sends HTML-formatted update reports via SMTP
- 📝 **Configurable Logging** - Automatic log rotation based on file size
- 🔧 **XML Configuration** - Easy configuration without modifying script code
- 🚫 **Flexible Reporting** - Option to disable reports for specific months
- 🎨 **Professional HTML Output** - Styled HTML tables with update details

## Prerequisites

### Required Modules

The script requires the following PowerShell modules:

1. **ConfigurationManager** (MECM Console)
   - Automatically loaded from SMS Admin UI path
   - Required for accessing MECM data

2. **Send-MailKitMessage**
   ```powershell
   Install-Module send-mailkitmessage
   ```
   - GitHub: https://github.com/austineric/Send-MailKitMessage

3. **PSWriteHTML**
   ```powershell
   Install-Module PSWriteHTML
   ```
   - GitHub: https://github.com/EvotecIT/PSWriteHTML

4. **PatchManagementSupportTools**
   ```powershell
   Install-Module PatchManagementSupportTools
   ```
   - GitHub: https://github.com/DambergC/PatchManagement/tree/main/PatchManagementSupportTools
   - Created by Christian Damberg, Cygate

### System Requirements

- Windows Server with MECM Site Server or Console installed
- PowerShell 5.1 or later
- Network access to MECM site server
- SMTP server access for sending emails
- Appropriate MECM permissions to query update information

## Installation

1. **Clone or download the script** to your MECM site server or console machine

2. **Create the XML configuration file** named `ScriptConfigPatchReleased.xml` in the same directory as the script

3. **Install required PowerShell modules** (see Prerequisites section)

4. **Configure the XML file** with your environment settings (see Configuration section)

5. **(Optional) Create a scheduled task** to run the script automatically

## XML Configuration File

The script requires a configuration file named `ScriptConfigPatchReleased.xml` in the same directory.

### Complete XML Template

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
	<SiteServer>Siteserver</SiteServer> ## smsprovider
	<Mailfrom>no-reply@domain.org</Mailfrom>
	<MailSMTP>smtp.domain.org</MailSMTP>
	<MailPort>25</MailPort>
	<MailCustomer>My company</MailCustomer>
</Configuration>
```

### XML Configuration Parameters

#### Site Server Configuration

| Parameter | Description | Example |
|-----------|-------------|---------|
| `SiteServer` | FQDN of your MECM site server | `SITESERVER.contoso.com` |

#### Email Settings

| Parameter | Description | Example |
|-----------|-------------|---------|
| `MailSMTP` | SMTP server address | `smtp.office365.com` |
| `MailPort` | SMTP port number | `25`, `587`, or `465` |
| `Mailfrom` | Sender email address | `noreply@company.com` |
| `MailCustomer` | Customer/company name for email subject | `Contoso Inc` |

#### Recipients Configuration

Add multiple `<Recipients>` blocks for each email recipient:

```xml
<Recipients>
  <Recipients>
    <Email>recipient@company.com</Email>
  </Recipients>
</Recipients>
```

#### Logging Configuration

| Parameter | Description | Example |
|-----------|-------------|---------|
| `Path` | Directory path for log files | `C:\Logs\` |
| `Name` | Log file name | `windowsupdatedeployed.log` |
| `Logfilethreshold` | Max log size in bytes before rotation | `5242880` (5 MB) |

#### Update Deployment Settings

| Parameter | Description | Example |
|-----------|-------------|---------|
| `LimitDays` | How many days back to search for updates | `-30` (last 30 days) |
| `DaysAfterPatchToRun` | Days after Patch Tuesday to send report | `7` (one week after) |
| `UpdateGroupName` | MECM Software Update Group name pattern | `2024-* - All Software Updates` |

#### Disable Report Months

Specify months (1-12) when reports should NOT be sent:

```xml
<DisableReportMonth>
  <DisableReportMonth>
    <Number>7</Number>  <!-- Skip July -->
  </DisableReportMonth>
  <DisableReportMonth>
    <Number>12</Number> <!-- Skip December -->
  </DisableReportMonth>
</DisableReportMonth>
```

## Usage

### Manual Execution

Run the script manually from the script directory:

```powershell
cd C:\Path\To\Script
.\Send-WindowsUpdateDeployed.ps1
```

### Scheduled Task

Create a scheduled task to run automatically:

1. Open Task Scheduler
2. Create a new task with the following settings:
   - **Trigger**: Daily at a specific time (e.g., 8:00 AM)
   - **Action**: Start a program
     - Program: `PowerShell.exe`
     - Arguments: `-ExecutionPolicy Bypass -File "C:\Path\To\Send-WindowsUpdateDeployed.ps1"`
     - Start in: `C:\Path\To\Script\Directory`
   - **Run with highest privileges**: Enabled
   - **Run whether user is logged on or not**: Enabled

### Script Behavior

The script will:

1. ✅ Check if the current month is in the disabled months list - **Exit if disabled**
2. ✅ Calculate Patch Tuesday for the current month
3. ✅ Determine if today equals Patch Tuesday + configured days
4. ✅ If today matches:
   - Query MECM for updates in the specified Software Update Group
   - Generate HTML report with update details
   - Send email to configured recipients
5. ✅ Log all activities to the log file
6. ✅ Rotate log file if it exceeds the size threshold

## Email Report Content

The generated email includes:

### Email Subject
```
WindowsUpdate [MailCustomer] [Month] [Year]
```
Example: `WindowsUpdate Contoso Inc January 2024`

### Email Body

The HTML email contains a styled table with the following columns:

- **ArticleID** - Microsoft KB article number
- **Title** - Update display name
- **LocalizedDescription** - Detailed description
- **DatePosted** - When Microsoft released the update
- **Deployed** - Whether the update is deployed
- **URL** - Microsoft support article link
- **Severity** - Update severity (Critical, Important, etc.)

### Example Email Preview

```
Windows updates January 2024

The following updates from Microsoft are deployed this month.

[Styled HTML table with update information]

Report created 1/15/2024 10:30:00 AM from SITESERVER
```

## Logging

### Log File Location

Logs are written to the path specified in the XML configuration:
```
[Logfile.Path]\[Logfile.Name]
```
Example: `C:\Logs\windowsupdatedeployed.log`

### Log Entry Format

```
yyyy/MM/dd HH:mm:ss [Log Message]
```

### Log Rotation

When the log file exceeds the `Logfilethreshold`:
1. Current log is renamed with timestamp: `windowsupdatedeployed_YYYY-MM-DD-HHMM.log`
2. Moved to `OLDLOG` subdirectory
3. New log file is created
4. `OLDLOG` directory is created automatically if it doesn't exist

### Sample Log Output

```
2024/01/15 08:00:15 ======================= Send-WindowsUpdateDeployed.ps1 - Script START =============================
2024/01/15 08:00:16 =====================Processing Software Update Group 2024-* - All Software Updates==========================
2024/01/15 08:00:45 ========================== Send-WindowsUpdateDeployed.ps1 - Mail on it's way to admin1@company.com, admin2@company.com
2024/01/15 08:00:46 ========================== Send-WindowsUpdateDeployed.ps1 - Script exit! ==========================
```

## Functions

### Rotate-Log
Checks log file size and rotates if threshold is exceeded.

### Write-Log
Writes timestamped entries to the log file.

### Get-ISO8601Week
Calculates ISO 8601 week numbers (used for scheduling logic).

### Get-CMSiteCode
Retrieves the MECM site code from WMI.

## Troubleshooting

### Script doesn't run on expected date

**Check:**
- Verify Patch Tuesday calculation: The script uses `Get-PatchTuesday` function
- Confirm `DaysAfterPatchToRun` setting in XML
- Review log file for date comparison messages

**Example log message:**
```
Date not equal patchtuesday 01/09/2024 and its now 01/08/2024. This report will run 01/16/2024
```

### No email received

**Check:**
1. SMTP server settings in XML
2. Network connectivity to SMTP server
3. Sender and recipient email addresses
4. SMTP server logs for rejected messages
5. Firewall rules for SMTP port

### Module not found errors

**Solution:**
```powershell
# Install missing modules
Install-Module send-mailkitmessage -Scope CurrentUser
Install-Module PSWriteHTML -Scope CurrentUser
Install-Module PatchManagementSupportTools -Scope CurrentUser
```

### MECM connection issues

**Check:**
1. Script is running on machine with MECM console installed
2. `SMS_ADMIN_UI_PATH` environment variable is set
3. User account has MECM permissions
4. Site server is accessible

### Updates not appearing in report

**Check:**
1. `UpdateGroupName` pattern matches your Software Update Groups
2. `LimitDays` setting covers the period you want to report
3. Updates exist in MECM and match the criteria

## Security Considerations

### Credentials

The script currently has placeholders for SMTP credentials:
```powershell
$Credential=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))
```

**For production use:**
- Store credentials securely using `Export-Clixml`
- Use Windows Integrated Authentication if possible
- Avoid hardcoding passwords in the script

### Permissions

The script requires:
- **Read access** to MECM Software Updates
- **SMTP send** permissions
- **Write access** to log file directory

## Customization

### Modify HTML Styling

Edit the `$header` variable to customize email appearance:
```powershell
$header = @"
<style>
    th {
        background-color: #YourColor;
        /* Add your styles */
    }
</style>
"@
```

### Add Company Logo

Replace the base64 image in the `$pre` variable:
```powershell
$pre = @"
<IMG SRC="data:image/jpg;base64,YOUR_BASE64_ENCODED_IMAGE" ALT="logo">
"@
```

### Change Report Criteria

Modify the filtering logic:
```powershell
$UpdatesFound = $result | Where-Object { 
    $_.dateposted -ge $limit -and 
    $_.Severity -eq "Critical" 
}
```

## Version History

- **10/16/2023** - Initial creation
- **03/25/2024** - Updated version (4:00 PM)

## Author

**Christian Damberg**  
Telia Cygate AB

## Related Projects

- [PatchManagement](https://github.com/DambergC/PatchManagement)
- [PatchManagementSupportTools](https://github.com/DambergC/PatchManagement/tree/main/PatchManagementSupportTools)

## License

Please refer to the repository license file.

## Support

For issues and questions:
1. Review the troubleshooting section
2. Check log files for detailed error messages
3. Verify XML configuration
4. Ensure all prerequisites are met

## Additional Resources

- [Microsoft Patch Tuesday Information](https://docs.microsoft.com/en-us/security-updates/)
- [MECM Software Updates Documentation](https://docs.microsoft.com/en-us/mem/configmgr/sum/)
- [ISO 8601 Week Date](https://en.wikipedia.org/wiki/ISO_week_date)
