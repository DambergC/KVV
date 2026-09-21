# Update Bulletin Report

This folder contains a scheduled PowerShell report that queries a SQL database for recent update bulletins, filters out unwanted entries, builds an HTML summary, and sends it by email.

## Files

- `UpdateBulletinReport.ps1` — main script that loads configuration, queries the database, filters results, creates HTML, and sends the report.
- `UpdateBulletinReport.config.xml` — environment-specific settings for database connection, filtering, logging, and email delivery.
- `readme.md` — documentation for the script and its configuration.

## Purpose of the script

The script is designed to:

- connect to a SQL Server instance
- read recent records from the `v_UpdateInfo` view
- filter out expired or superseded records
- keep only items in the configured date window (`DaysBack`)
- exclude bulletin titles/descriptions that match configured patterns
- generate a static HTML email body with a summary table
- send the result to one or more recipients through SMTP
- log execution details to a log file

This is a good fit for a scheduled task or Windows Task Scheduler job that emits a daily or periodic update bulletin mail.

## Script flow

The PowerShell script follows this sequence:

1. Read the XML configuration file.
2. Validate required settings such as SQL instance, database name, SMTP server, and recipients.
3. Load the required modules.
4. Query the SQL database with `Invoke-DbaQuery`.
5. Apply exclusion filters to titles and descriptions.
6. Generate a temporary HTML file in `%TEMP%`.
7. Build a styled email body with:
   - Date posted
   - Article ID
   - Severity
   - Title
   - Description
   - Info URL
8. Send the email by calling `Send-MailKitMessage`.
9. Clean up the temporary HTML unless `-KeepHtmlFile` is used.

## Required modules

The script explicitly checks for and imports two PowerShell modules:

### 1) dbatools

Required for database access and SQL helper cmdlets such as:

- `Invoke-DbaQuery`
- `Set-DbatoolsInsecureConnection`

The script expects `dbatools` to be installed for the user account running the scheduled task.

Example installation:

```powershell
Install-Module dbatools -Scope CurrentUser
```

### 2) Send-MailKitMessage

Required for SMTP email sending using the MailKit library.

Example installation:

```powershell
Install-Module Send-MailKitMessage -Scope CurrentUser
```

If these modules are missing, the script stops with an error and logs the failure.

## Configuration file

The script reads settings from `UpdateBulletinReport.config.xml`.

Example structure:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <Database>
    <SqlInstance>servername</SqlInstance>
    <DatabaseName>DB</DatabaseName>
  </Database>

  <Query>
    <DaysBack>1</DaysBack>
    <ExcludeTitleContains>
      <Pattern>25H2</Pattern>
      <Pattern>Defender</Pattern>
    </ExcludeTitleContains>
    <ExcludeDescriptionContains>
      <Pattern>Install the latest version of Windows</Pattern>
    </ExcludeDescriptionContains>
  </Query>

  <Logging>
    <LogPath>d:\Scripts\UpdateBulletinReport\LogFiles\UpdateBulletinReport.log</LogPath>
  </Logging>

  <Email>
    <SmtpServer>smtp.company.se</SmtpServer>
    <SmtpPort>25</SmtpPort>
    <From>noreply@company.se</From>
    <To>
      <Recipient>user@company.se</Recipient>
    </To>
    <Subject>Recent Software Update Bulletins</Subject>
    <HtmlLogoMimeType>image/png</HtmlLogoMimeType>
    <HtmlLogo><![CDATA[data:image/png;base64,...]]></HtmlLogo>
  </Email>
</Configuration>
```

## Settings explained

### Database settings

#### `Configuration/Database/SqlInstance`

The SQL Server instance name or host used to connect to the database.

Example:

```xml
<SqlInstance>servername</SqlInstance>
```

#### `Configuration/Database/DatabaseName`

The name of the target database.

Example:

```xml
<DatabaseName>DB</DatabaseName>
```

### Query settings

#### `Configuration/Query/DaysBack`

Number of days back to include when selecting update bulletins.

Example:

```xml
<DaysBack>1</DaysBack>
```

The query selects rows where `DatePosted` is within the last `N` days.

#### `Configuration/Query/ExcludeTitleContains/Pattern`

List of text values that should cause a bulletin to be excluded if they appear in the title.

Example values in the current config:

- `25H2`
- `26H1`
- `arm64`
- `Defender`
- `Apps`
- `Edge-Beta`
- `Edge-Dev`
- `x86`
- `Office 2019`
- `Office LTSC 2021`

These patterns are evaluated with case-insensitive substring matching.

#### `Configuration/Query/ExcludeDescriptionContains/Pattern`

List of text values that should cause exclusion if they appear in the description.

Example:

```xml
<Pattern>Install the latest version of Windows</Pattern>
```

This helps suppress less relevant or boilerplate update notices.

### Logging settings

#### `Configuration/Logging/LogPath`

Path to the log file that records script execution and errors.

Example:

```xml
<LogPath>d:\Scripts\UpdateBulletinReport\LogFiles\UpdateBulletinReport.log</LogPath>
```

The log file is created automatically if the folder does not exist.

### Email settings

#### `Configuration/Email/SmtpServer`

SMTP host used to send the report.

#### `Configuration/Email/SmtpPort`

SMTP port number.

Example:

```xml
<SmtpPort>25</SmtpPort>
```

#### `Configuration/Email/From`

Sender address used in the email header.

#### `Configuration/Email/To/Recipient`

One or more recipients. The XML supports multiple recipient entries.

Example:

```xml
<To>
  <Recipient>user1@company.se</Recipient>
  <Recipient>user2@company.se</Recipient>
</To>
```

#### `Configuration/Email/Subject`

Base subject of the email. The script appends the date range automatically, for example:

- `Recent Software Update Bulletins - 2026-09-01 to 2026-09-15`

#### `Configuration/Email/HtmlLogoMimeType`

Mime type for the embedded logo.

Supported examples:

- `image/png`
- `image/jpeg`
- `image/gif`

#### `Configuration/Email/HtmlLogo`

Logo used in the HTML email. This may be:

- a complete data URI (`data:image/png;base64,...`)
- a raw base64 string
- an HTTP/HTTPS URL

The script will generate an HTML `<img>` tag using the configured logo when valid.

## Severity handling

The script translates database severity values into readable labels for the HTML report:

- `10` = Critical
- `8` = Important
- `6` = Moderate
- `2` = Low
- `0` = Unspecified

The report also includes CVSS range text and a short description for each severity level.

## HTML output

The generated email contains a table with columns:

- Date posted
- Article ID
- Severity
- Title
- Description
- Info URL

If no bulletins are found in the selected date range, the email contains a simple message:

> No new bulletins were found in the last X day(s).

## Script parameters

The script supports the following parameters:

```powershell
UpdateBulletinReport.ps1 [-ConfigPath <string>] [-KeepHtmlFile]
```

### `-ConfigPath`

Optional path to the XML configuration file. Default:

```powershell
D:\Scripts\UpdateBulletinReport\UpdateBulletinReport.config.xml
```

### `-KeepHtmlFile`

If specified, the generated temporary HTML file is not deleted after email sending.

## Operational notes

- The script writes detailed logs to the configured file path.
- It is designed to run as a scheduled task without interactive console requirements.
- It treats missing required configuration values as fatal errors.
- It also validates that the SMTP port is numeric and in a valid range.
- The script uses `Set-StrictMode -Version Latest` and `$ErrorActionPreference = "Stop"` for stricter execution safety.

## Example use

From PowerShell:

```powershell
.\UpdateBulletinReport.ps1
```

Or with a custom config file:

```powershell
.\UpdateBulletinReport.ps1 -ConfigPath "D:\Scripts\UpdateBulletinReport\Custom.config.xml"
```

And to preserve the generated HTML file:

```powershell
.\UpdateBulletinReport.ps1 -KeepHtmlFile
```

## Summary

This script is a lightweight scheduled reporting tool for Microsoft update bulletin data. It collects relevant bulletin records, filters them according to business rules, formats them as a professional HTML email, and delivers the result to stakeholders.

The main values you typically need to customize are:

- SQL server and database
- reporting time window
- exclusion keywords
- log path
- SMTP details
- recipients
- optional company logo

These are all centralized in the XML configuration file.
