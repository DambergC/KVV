# Device Maintenance Windows Report (MECM/SCCM)

Generates an HTML + CSV report of **MECM (Microsoft Endpoint Configuration Manager / SCCM)** device Maintenance Windows and emails the result to administrators.

Script: `DeviceMaintenanceWindows/Send-MMCM_DeviceMaintenanceWindows.ps1`

---

## What the script does (high level)

When the script runs (typically via **Scheduled Task** on the MECM site server), it:

1. Loads configuration from `ScriptConfigMW.xml`.
2. Connects to the MECM site (determines SiteCode via WMI and switches to the `SiteCode:` PSDrive).
3. Calculates Patch Tuesday dates for the current and next month.
4. On the configured “report day” (Patch Tuesday + offset days):
   - Reads all devices in a target collection (`$collectionidToCheck` in the script).
   - For each device:
     - Finds all collections the device is a member of.
     - Filters to collections that have Maintenance Windows.
     - Retrieves Maintenance Windows and selects relevant ones.
     - Enriches each result row by querying an external SQL database for an application name.
   - Exports the result to CSV.
   - Generates an HTML report (PSWriteHTML).
   - Sends an email (Send-MailKitMessage) with HTML body and attaches the HTML + CSV files.

If it is **not** the report day, the script logs and exits without creating/sending the report.

---

## Outputs

The script writes:

- **HTML report file** (PSWriteHTML)
- **CSV report file**
- **Log files**

> Note: In the script, paths are currently hard-coded to `G:\Scripts\...` and filenames include the date. The XML also contains paths (see below), but the posted script currently uses its own variables for log/out paths.

---

## Prerequisites

### Runtime environment

- Windows Server where the MECM admin console / provider access is available (commonly the **site server**).
- PowerShell 5.1 (typical for MECM environments) or PowerShell 7+ *if* all required modules support it.
- Network access from the running host to:
  - MECM SMS Provider / Site Server WMI
  - SQL server used for enrichment (`$dbserver` in the script)
  - SMTP server

### Permissions

The account running the Scheduled Task typically needs:

- MECM rights to read:
  - Collection members
  - Device collection memberships
  - Maintenance windows
  - Collection metadata
- WMI permission to query on the site server namespace:
  - `root\SMS`, class `SMS_ProviderLocation`
- SQL permissions for the enrichment query (SELECT against the external database/table)
- Write access to output/log folders (for example `G:\Scripts\Outfiles\` and `G:\Scripts\Logfiles\`)
- Permission to send mail via the configured SMTP server

### PowerShell modules / components

The script imports/uses these modules:

- **Send-MailKitMessage**
  - Used to send email and attach the report files.
- **PSWriteHTML**
  - Used to generate the HTML report page.
- **PatchManagementSupportTools**
  - Used for helper functions (for example `Get-PatchTuesday`, and likely `Get-CMModule`).

Additionally:

- **ConfigurationManager** module (MECM/SCCM PowerShell cmdlets)
  - The script uses cmdlets such as `Get-CMCollectionMember`, `Get-CMMaintenanceWindow`, `Get-CMCollection`, etc.
- **SqlServer** module (for `Invoke-Sqlcmd`)
  - The script uses `Invoke-Sqlcmd` to enrich results with application information.

---

## Configuration: `ScriptConfigMW.xml`

The script loads XML like:

- `.\ScriptConfigMW.xml` (relative path)

### Provided XML example

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
    <Recipients email="christian.damberg@kriminalvarden.se"/>
    <Recipients email="Joakim.Stenqvist@kriminalvarden.se"/>
    <Recipients email="Christian.Brask@kriminalvarden.se"/>
    <Recipients email="Magnus.Jonsson6@kriminalvarden.se"/>
    <Recipients email ="Keiarash.Naderifarsani@kriminalvarden.se"/>
    <Recipients email="lars.garlin@kriminalvarden.se"/>
    <Recipients email="sockv@kriminalvarden.se"/>
    <Recipients email="Hans.Pettersson@kriminalvarden.se"/>
    <Recipients email="Magnus.Eklof@kriminalvarden.se"/>
    <Recipients email="Fredrik.Alderin@kriminalvarden.se"/>
    <Recipients email="Peter.Overhem@kriminalvarden.se"/>
    <Recipients email="Nicklas.Tigerberg@kriminalvarden.se"/>
    <Recipients email="Jens.Wolinder@kriminalvarden.se"/>
    <Recipients email="Peter.Bystrom@kriminalvarden.se"/>
  </Recipients>

  <UpdateDeployed>
    <LimitDays>-25</LimitDays>
    <UpdateGroupName>Server Patch Tuesday</UpdateGroupName>
    <DaysAfterPatchToRun>1</DaysAfterPatchToRun>
  </UpdateDeployed>

  <SiteServer>vntapp0780</SiteServer>
  <Mailfrom>no-reply@kvv.se</Mailfrom>
  <MailSMTP>smtp.kvv.se</MailSMTP>
  <MailPort>25</MailPort>
  <MailCustomer>Kriminalvarden - IT</MailCustomer>
</Configuration>
```

### How the script uses the XML (current behavior)

The posted script reads these values:

- `Configuration/SiteServer`
  - Used for WMI query to determine the MECM SiteCode.
- `Configuration/MailSMTP`
  - SMTP server for Send-MailKitMessage.
- `Configuration/Mailfrom`
  - From address for the email.
- `Configuration/MailPort`
  - SMTP port.
- `Configuration/MailCustomer`
  - Included in the subject line.
- `Configuration/Recipients`
  - Used to build the recipient list.

> Important: the **recipient parsing in the posted script expects a different XML shape** (it looks for `.Recipients.Email` as elements), while your XML stores recipients as attributes (`email="..."`).
>
> That means: **with the script exactly as posted, recipient extraction may return empty** unless the script is adjusted to read the `email` attribute.

### XML fields present but not used by the posted script

Your XML also contains:

- `Configuration/Logfile/*`
- `Configuration/HTMLfilePath`
- `Configuration/RunScript/Job[@DeploymentID,@Offsetdays,@Description]`
- `Configuration/DisableReportMonth/*`
- `Configuration/UpdateDeployed/*`

These look like they are shared config for other patch/report scripts, or for an extended version of this script. The posted `Send-MMCM_DeviceMaintenanceWindows.ps1` does not appear to consume these nodes (it uses hard-coded paths and variables for some of the same concepts).

---

## Script flow (detailed)

### 1) Initialization / configuration

- Sets script name/version.
- Loads `ScriptConfigMW.xml`.
- Defines site server and mail settings (mostly from XML).
- Defines output paths and log paths (hard-coded in script).
- Imports required modules.

### 2) Connect to MECM site

- Determines SiteCode using WMI query against the SiteServer:
  - `root\SMS : SMS_ProviderLocation`
- Calls `Get-CMModule` (likely a helper that imports ConfigurationManager module).
- Switches location to the MECM PSDrive: `<SiteCode>:`.

### 3) Patch Tuesday / date calculations

- Calculates Patch Tuesday for current month and next month via `Get-PatchTuesday`.
- Defines the day the report should run:
  - `ReportdayCompare = PatchTuesdayThisMonth + DaysAfterPatchTuesdayToReport`

> Note: In the posted script `$DaysAfterPatchTuesdayToReport` is set to `-6`, which means “**6 days before Patch Tuesday**”, despite the variable name saying “After”.

### 4) Exit early if month is disabled

- Checks if current month is in `$DisableReport`.
- If disabled, logs and exits.

> Note: In the posted script `$DisableReport` is an empty string. For this to work properly, it should be an array of month numbers, for example:
> `@(7, 12)` to skip July and December.

### 5) Collect device maintenance window data (only on report day)

If today equals the report day:

- Reads all devices from a specific collection:
  - `Get-CMCollectionMember -CollectionId $collectionidToCheck`

For each device:

1. Finds all collections the device belongs to:
   - `Get-CMClientDeviceCollectionMembership -ComputerName <device> -SiteServer <siteserver> -SiteCode <sitecode>`
2. Filters to collections where `ServiceWindowsCount -gt 0`
3. For each such collection:
   - Loads MWs with `Get-CMMaintenanceWindow -CollectionId <collectionId>`
   - For each MW:
     - If recurrence type is “1” (one recurrence type), only include MWs with StartTime between Patch Tuesday this month and Patch Tuesday next month.
     - If recurrence type is “3”, include (no date filtering in posted script).
4. Enriches each output row by querying an external SQL database:
   - `Invoke-Sqlcmd -ServerInstance $dbserver -Database serverlista -Query <SELECT ...>`
5. Produces result objects with fields:
   - `Applikation`
   - `Server`
   - `Startdatum`
   - `Starttid`
   - `Varaktighet`
   - `Deployment`

Finally:

- Exports results to CSV.
- Builds an HTML report page (PSWriteHTML).

If today is not the report day:

- Logs “Date not equal … this report will run …”
- Exits.

### 6) Email report

- Creates a static HTML email body (does not embed the table).
- Builds a recipient list from XML.
- Attaches the generated HTML + CSV.
- Sends email via `Send-MailKitMessage`.

---

## Scheduling (recommended)

Create a Windows Scheduled Task on the MECM site server:

- Trigger: Daily (or a cadence you prefer)
- Action: Start a program:
  - `powershell.exe -ExecutionPolicy Bypass -File "<path>\Send-MMCM_DeviceMaintenanceWindows.ps1"`
- Start in (important): The folder containing the script and `ScriptConfigMW.xml`
- Run whether user is logged on or not
- Use a service account with the permissions listed above

Because the script self-checks the date (Patch Tuesday offset), it is safe to run daily.

---

## Troubleshooting

### Email recipients are empty / mail not sent
Your XML stores recipients as attributes (`email="..."`), but the posted script appears to read them as elements.

Symptoms:
- `$RecipientList` ends up empty.
- Send-MailKitMessage may fail or send to nobody.

Fix:
- Update recipient parsing to read the `email` attribute.

### `Invoke-Sqlcmd` not found
Install/import the SqlServer module:
- Install: `Install-Module SqlServer`
- Import: `Import-Module SqlServer`

### MECM cmdlets not found / SiteCode drive missing
- Ensure Configuration Manager console / module is installed.
- Ensure `Get-CMModule` properly imports ConfigurationManager.
- Verify WMI query to `SMS_ProviderLocation` works and returns a SiteCode.

### Report never runs
- Check `$DaysAfterPatchTuesdayToReport` (negative values run before Patch Tuesday).
- Check locale issues with `ToShortDateString()` comparison.
- Check logs for “Date not equal … will run …”.

---

## Notes / known mismatches between XML and script

- The XML contains `LogfilePath`, `HTMLfilePath`, thresholds, and job definitions, but the posted script uses hardcoded file paths and does not reference those XML nodes.
- The XML recipients use `email="..."` attributes; the script’s parsing method should be aligned with that structure.
- Subject uses `$monthname` and `$year` in the posted script snippet; ensure these variables are defined in the script or adjust the subject.

---

## License / Disclaimer

All scripts are offered **AS IS** with no warranty. Test in a non-production environment before using in production.
