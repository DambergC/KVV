# Send-WindowsUpdateDeployed

A PowerShell automation script for reporting Windows Update deployments through Microsoft Endpoint Configuration Manager (MECM/SCCM). The script generates and emails HTML reports of deployed updates.

## Overview

The script reads its settings from an XML configuration file and then:

- Queries MECM for software updates in a specified Software Update Group
- Builds a styled HTML report (PSWriteHTML)
- Emails the report via SMTP (Send-MailKitMessage)

It is designed to run on a MECM site server/console machine, either manually or as a Scheduled Task.

## New / Updated Features (Script + XML)

### Configuration file v2 (required)

The PowerShell script currently expects an XML file named:

- `Send-WindowsUpdateDeployedv2.xml` (must be located in the same folder as `Send-WindowsUpdateDeployed.ps1`)

> Note: The repository also contains `Send-WindowsUpdateDeployed.XML` as an example/template, but the script is hard-coded to load `Send-WindowsUpdateDeployedv2.xml`.

### XML parsing helpers (PowerShell 5.1 compatible)

The script includes small helper functions to reliably read XML values:

- `Get-XmlText` (trimmed string values)
- `Get-XmlInt` (integers with defaults)
- `Get-XmlIntArray` (list of integers)

This makes the configuration more robust to missing/empty values.

### Recipient schema compatibility (new)

The script supports **two** recipient schemas:

1) **Nested element** format:

```xml
<Recipients>
  <Recipients>
    <Email>recipient@company.com</Email>
  </Recipients>
</Recipients>
```

2) **Attribute** format (used by the included XML example):

```xml
<Recipients>
  <Recipients email="recipient@company.com" />
</Recipients>
```

Recipients are de-duplicated automatically.

### DisableReportMonth schema compatibility (new)

The script supports **two** ways to define disabled months:

1) Element format:

```xml
<DisableReportMonth>
  <DisableReportMonth>
    <Number>7</Number>
  </DisableReportMonth>
</DisableReportMonth>
```

2) Attribute format (used by the included XML example):

```xml
<DisableReportMonth>
  <DisableReportMonth Number="7" />
</DisableReportMonth>
```

If the current month is disabled, the script exits early.

### Built-in configuration validation (new)

The script validates required settings at startup and fails fast with clear error messages if any are missing:

- `SiteServer`
- `MailSMTP`
- `Mailfrom`
- `MailCustomer`
- `UpdateDeployed/UpdateGroupName`
- At least one recipient

### Optional logging (new)

Logging is optional. If `Logfile/Path` and `Logfile/Name` are present, the script writes timestamped log entries.

- Log format: `yyyy/MM/dd HH:mm:ss <message>`

### Module auto-check + import (new)

The script ensures required modules are installed and imports them:

- `PSWriteHTML`
- `Send-MailKitMessage`

If a module is missing, the script stops with a clear message.

### Auto-detect MECM CMSite PSDrive (new)

Instead of requiring a hard-coded site code, the script automatically discovers the first `CMSite` PSDrive and temporarily switches into it to run:

- `Get-CMSoftwareUpdate -Fast -UpdateGroupName <name>`

This improves portability across MECM environments.

### Updated data collection logic (new)

The script builds the report dataset by selecting common update properties from MECM:

- `ArticleID`
- `LocalizedDisplayName` (shown as `Title`)
- `LocalizedDescription`
- `DatePosted`
- `IsDeployed` (shown as `Deployed`)
- `LocalizedInformativeURL` (shown as `URL`)
- `SeverityName` (shown as `Severity`)

### HTML report redesigned to match DPStatus style (new)

The report is generated using **PSWriteHTML** and includes:

- Embedded base64 image/logo
- A "Summary" section showing:
  - Date (`yyyy-MM-dd`)
  - SiteServer
  - UpdateGroupName
  - Number of updates returned
  - `LimitDays`
  - `DaysAfterPatchTuesdayToReport`
- A table of updates (or "Inga poster." if empty)
- A footer with creation timestamp and computer name

### Email sending updated to Send-MailKitMessage + MimeKit (new)

The script builds a `MimeKit.InternetAddressList` for recipients and sends the HTML body using `Send-MailKitMessage`.

Subject format:

- `<MailCustomer> - Windows Updates <MonthName> <Year> - <yyyy-MM-dd>`

## Prerequisites

### Required Modules

Install the modules (on the machine that runs the scheduled task/script):

```powershell
Install-Module PSWriteHTML
Install-Module send-mailkitmessage
```

### System Requirements

- Windows Server with MECM Site Server or Console installed
- PowerShell 5.1 or later
- Network access to MECM site server
- SMTP server access for sending emails
- Appropriate MECM permissions to query update information

## Installation

1. Clone/download the folder to the MECM site server / console machine.
2. Create `Send-WindowsUpdateDeployedv2.xml` in the same directory as `Send-WindowsUpdateDeployed.ps1`.
3. Install the required PowerShell modules.
4. Configure the XML file.
5. (Optional) Create a scheduled task to run the script.

## XML Configuration File

The repository includes `Send-WindowsUpdateDeployed.XML` as an example. The script expects `Send-WindowsUpdateDeployedv2.xml`.

### Example XML (attribute-style recipients + disabled months)

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <Logfile>
    <Path>G:\Scripts\Logfiles\</Path>
    <Name>WindowsUpdateScript.log</Name>
    <Logfilethreshold>2000000</Logfilethreshold>
  </Logfile>

  <HTMLfilePath>G:\Scripts\OutFiles\</HTMLfilePath>

  <DisableReportMonth>
    <DisableReportMonth Number="7" />
    <DisableReportMonth Number="12" />
  </DisableReportMonth>

  <Recipients>
    <Recipients email="admin1@company.com" />
    <Recipients email="admin2@company.com" />
  </Recipients>

  <UpdateDeployed>
    <LimitDays>-25</LimitDays>
    <UpdateGroupName>Server Patch Tuesday</UpdateGroupName>
    <DaysAfterPatchToRun>9</DaysAfterPatchToRun>
  </UpdateDeployed>

  <SiteServer>Siteserver</SiteServer>
  <Mailfrom>no-reply@company.com</Mailfrom>
  <MailSMTP>smtp.company.com</MailSMTP>
  <MailPort>25</MailPort>
  <MailCustomer>My company</MailCustomer>
</Configuration>
```

## Usage

### Manual Execution

```powershell
cd C:\Path\To\Script-WindowsUpdateDeployed
.\Send-WindowsUpdateDeployed.ps1
```

### Scheduled Task

- Program: `PowerShell.exe`
- Arguments: `-ExecutionPolicy Bypass -File "C:\Path\To\Script-WindowsUpdateDeployed\Send-WindowsUpdateDeployed.ps1"`
- Start in: `C:\Path\To\Script-WindowsUpdateDeployed`
- Run with highest privileges: Enabled

## Troubleshooting

### Get-CMSoftwareUpdate not available

- Ensure the MECM console is installed
- Import the `ConfigurationManager` module and confirm a `CMSite` PSDrive exists

### No recipients found

Ensure you are using either:

- `/Configuration/Recipients/Recipients/Email` nodes, or
- `/Configuration/Recipients/Recipients[@email]` attributes