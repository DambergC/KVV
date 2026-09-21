# Server Compliance Report

This PowerShell script checks whether servers in one or more Configuration Manager collections are compliant with a specified Software Update Group (SUG) and sends a summarized HTML email report per collection.

The script is intended to run as a scheduled task on a daily basis. It evaluates server compliance relative to the current or most recent Patch Tuesday and sends emails only for collections configured for that day.

## Files

- `ServerComplianceReport.ps1` – main report script
- `ServerComplianceReport.config.xml` – XML configuration for database, collections, logging, and email recipients
- `Readme.md` – documentation for the script

## What the script does

- Reads the XML configuration file
- Determines the current Patch Tuesday and calculates the configured offset from that date
- Queries the ConfigMgr database for collection members and compliance status
- Identifies:
  - compliant servers
  - non-compliant servers
  - servers with unknown compliance state
- Builds an HTML report per collection
- Sends an email for each matching collection using SMTP
- Optionally keeps generated HTML files in `%TEMP%` when using the `-KeepHtmlFile` switch

## Requirements

The script requires the following PowerShell modules to be installed for the account running the scheduled task:

```powershell
Install-Module dbatools -Scope CurrentUser
Install-Module Send-MailKitMessage -Scope CurrentUser
```

The script also expects access to the SQL database used by Configuration Manager and a valid SMTP server.

## Configuration

The script uses the XML file `ServerComplianceReport.config.xml`.

Example structure:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <Database>
    <SqlInstance>servername</SqlInstance>
    <DatabaseName>DB</DatabaseName>
  </Database>

  <Query>
    <Collections>
      <Collection>
        <Name>SU - Server Updates - Server 100</Name>
        <Title>Server Patch Tuesday</Title>
        <XDays>10</XDays>
        <DaysAfterPatchTuesday>6</DaysAfterPatchTuesday>
      </Collection>
    </Collections>
  </Query>

  <Logging>
    <LogPath>d:\Scripts\ServerComplianceReport\LogFiles\ServerComplianceReport.log</LogPath>
  </Logging>

  <Email>
    <SmtpServer>smtp.company.se</SmtpServer>
    <SmtpPort>25</SmtpPort>
    <From>noreply@company.se</From>
    <To>
      <Recipient>user@company.se</Recipient>
    </To>
    <Subject>Server Patch Compliance Report</Subject>
    <HtmlLogoMimeType>image/png</HtmlLogoMimeType>
    <HtmlLogo><![CDATA[data:image/png;base64,...]]></HtmlLogo>
  </Email>
</Configuration>
```

### Database settings

- `SqlInstance` – SQL Server/instance name used by dbatools
- `DatabaseName` – database containing the ConfigMgr data

### Collection settings

Each `<Collection>` entry defines a collection and reporting schedule:

- `Name` – ConfigMgr collection name
- `Title` – Software Update Group title to check against
- `XDays` – maximum age in days for a server to be considered healthy/online
- `DaysAfterPatchTuesday` – how many days after Patch Tuesday this collection should run

The script checks the current date against the configured offset and only runs the collections matching that value.

### Logging settings

- `LogPath` – path to the log file used by the script

If the directory does not exist, the script creates it automatically.

### Email settings

- `SmtpServer` – SMTP server hostname
- `SmtpPort` – SMTP port number
- `From` – sender email address
- `To/Recipient` – one or more recipient addresses
- `Subject` – subject prefix for generated emails
- `HtmlLogo` – optional Base64 image or URL for an email logo
- `HtmlLogoMimeType` – MIME type for the embedded logo

## Required modules

The script imports and validates the following modules before running:

- `dbatools`
- `Send-MailKitMessage`

The script checks for these modules in the scheduled task account context and throws an error if they are missing.

## Execution

Run the script from PowerShell:

```powershell
.\ServerComplianceReport.ps1
```

Optional parameters:

```powershell
.\ServerComplianceReport.ps1 -ConfigPath "D:\Scripts\ServerComplianceReport\ServerComplianceReport.config.xml"
.\ServerComplianceReport.ps1 -KeepHtmlFile
.\ServerComplianceReport.ps1 -AsOfDate (Get-Date)
```

## Notes

- The script is designed to be used with scheduled tasks.
- If no configured collection matches the current Patch Tuesday offset, the script exits without sending email.
- Query results are displayed in a summary table with:
  - total evaluated servers
  - compliant servers
  - non-compliant servers
  - unknown compliance servers
- Email reports are sent one per matching collection.

## Example use case

A typical environment may define multiple server collections with different Patch Tuesday offsets. For example:

- Collection A runs 6 days after Patch Tuesday
- Collection B runs 8 days after Patch Tuesday
- Collection C runs 15 days after Patch Tuesday

This allows a report to be sent at the correct time for each server group without manually running the script every day.

## Troubleshooting

Common issues:

- Module not installed for the scheduled task user
- XML file missing or malformed
- SMTP server not reachable
- No collection matches the calculated offset for the day
- ConfigMgr collection or SUG name does not match the configured values

The script writes detailed logging to the configured log path and uses `Write-Error`/`Write-Warning` when issues occur.
