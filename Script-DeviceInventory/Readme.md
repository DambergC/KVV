# Send-DeviceInventory.ps1

Collects device inventory from a ConfigMgr (SCCM) database, generates an HTML report (optionally CSV), and emails the results to configured recipients.

This README documents purpose, requirements, configuration, usage and troubleshooting for the Send-DeviceInventory.ps1 script found in this repository.

---

## Features

- Queries ConfigMgr (CM_KV1) for physical Windows 10 / Windows 11 machines.
- Aggregates hardware, OS, user, network and selected installed software information.
- Generates an HTML report using PSWriteHTML and can optionally generate a CSV.
- Sends the HTML (and CSV if requested) via SMTP using Send-MailKitMessage.
- Rotates and writes logs, with configurable log file path and basic rotation policy.
- Supports configurable SMTP authentication and secure-connection option.
- Provides switches to skip sending email and/or skip logging.

---

## Requirements

- Windows PowerShell (tested with PowerShell 5.1; should work in later Windows PowerShell).
- Connectivity and read permission to the ConfigMgr SQL instance (CM_KV1).
- The following PowerShell modules installed and available:
  - send-mailkitmessage
  - PSWriteHTML
  - PatchManagementSupportTools
  - SQLPS (or SqlServer module providing Invoke-Sqlcmd)
- Write permission to configured output and log directories.
- A configuration XML file (default: `.\ScriptConfigDevInvent.xml`) with required fields (see below).

---

## Files

- Send-DeviceInventory.ps1 — main script (this repository)
  - Permalink: https://github.com/DambergC/KVV/blob/acba86744d0ce2fea6a1157cde556f437aceda0e/Send-DeviceInventory.ps1

---

## Script Parameters

- `-ConfigPath <string>`  
  Path to configuration XML. Default: `.\ScriptConfigDevInvent.xml`

- `-NoEmail`  
  Switch. When present, the script will not send emails (report still generated).

- `-NoLog`  
  Switch. When present, the script will not write to log files (console output remains).

- `-GenerateCSV`  
  Switch. When present, the script will also export the inventory as a CSV file.

Example:
```powershell
.\Send-DeviceInventory.ps1 -ConfigPath "C:\Scripts\ScriptConfigDevInvent.xml" -GenerateCSV
```

---

## Configuration (XML)

The script reads an XML configuration. The script expects a `Configuration` element with keys used in the script. The minimal configuration must include:

- SiteServer — SQL Server instance (used by Invoke-Sqlcmd)
- MailSMTP — SMTP server address
- Mailfrom — From address for email
- MailPort — optional SMTP port (defaults to 25 if omitted)
- MailCustomer — optional string added to email subject
- Logfile
  - Name — filename for the log (e.g. DeviceInventory.log)
  - Path — path for log file directory (e.g. G:\Scripts\Logs)
- ScriptVersion — optional expected script version (the script will warn if mismatch)
- UseSecureConnection — optional: true/false (attempt secure connection if available)
- UseAuthentication — optional: true/false
- EmailUsername, EmailPassword — required if UseAuthentication is true
- Recipients — list or collection of recipients (support for several XML shapes)
- BCCRecipients — optional list or collection

Example configuration snippet:
```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <SiteServer>sqlserver.domain.local</SiteServer>
  <ScriptVersion>2.0</ScriptVersion>

  <MailSMTP>smtp.corp.domain.local</MailSMTP>
  <Mailfrom>no-reply@domain.local</Mailfrom>
  <MailPort>25</MailPort>
  <MailCustomer>KVV</MailCustomer>

  <UseSecureConnection>false</UseSecureConnection>
  <UseAuthentication>false</UseAuthentication>
  <EmailUsername></EmailUsername>
  <EmailPassword></EmailPassword>

  <Recipients>
    <!-- Accepts several structures. Example 1: multiple Email elements -->
    <Email>it-team@domain.local</Email>
    <Email>service@domain.local</Email>

    <!-- Example 2: nested collection
    <Recipients>
      <Recipient>
        <Email>it-team@domain.local</Email>
      </Recipient>
    </Recipients>
    -->
  </Recipients>

  <BCCRecipients>
    <Email>archive@domain.local</Email>
  </BCCRecipients>

  <Logfile>
    <Name>DeviceInventory.log</Name>
    <Path>G:\Scripts\Outfiles\Logs</Path>
  </Logfile>
</Configuration>
```

Notes:
- The script contains robust parsing to accept a few different shapes of Recipients/BCCRecipients (single string, multiple Email elements, nested Recipient collections). Validate the XML after creating/updating it.
- Storing plain passwords in XML is insecure. Prefer using secure storage (credential manager, service account, or encrypted files) where possible.

---

## Output

- HTML report: saved to path specified in the script as `$HTMLFileSavePath` (default built with date, e.g. `G:\Scripts\Outfiles\DeviceInvent_YYYYMMDD.HTML`).
- CSV report (optional with `-GenerateCSV`): `DeviceInvent_YYYYMMDD.csv`.
- Log file: as configured in XML (rotates if it exceeds configured threshold in the script).

---

## Running the Script

Basic run (default config path):
```powershell
.\Send-DeviceInventory.ps1
```

Run without sending email:
```powershell
.\Send-DeviceInventory.ps1 -NoEmail
```

Generate CSV and use specific config:
```powershell
.\Send-DeviceInventory.ps1 -ConfigPath "C:\Scripts\ScriptConfigDevInvent.xml" -GenerateCSV
```

Schedule via Task Scheduler:
- Create a scheduled task that runs PowerShell with an argument block:
```text
Program/script: powershell.exe
Add arguments: -ExecutionPolicy Bypass -NoProfile -File "C:\Scripts\Send-DeviceInventory.ps1" -ConfigPath "C:\Scripts\ScriptConfigDevInvent.xml" -GenerateCSV
```
- Ensure the scheduled task runs with an account that has:
  - Read access to configuration file.
  - Write access to output/log paths.
  - Permissions to query the ConfigMgr SQL database.

---

## Troubleshooting & Tips

- "Module not found" errors: install the missing modules (Install-Module Send-MailKitMessage, PSWriteHTML, etc.) or ensure they're available in the session.
- SQL errors: check `$SiteServer`, connectivity, and that the account used by the script has read access to the CM_KV1 database. The script uses `Invoke-Sqlcmd` — confirm SQLPS (or SqlServer) is available.
- No email sent: check SMTP server value, validate recipients in the config, and check log for warnings about missing recipients or SMTP server.
- Logs: If logs are not written, confirm the configured Logfile.Path exists and the executing account has write permissions, or run with `-NoLog` to avoid errors temporarily.
- Credentials: avoid plain text passwords in the XML. Use secure mechanisms where possible.
- If the HTML looks empty: confirm that the SQL query returned rows; the HTML generator expects `$resultColl` to have data.

---

## Customization

- SQL query: the script contains a comprehensive query inside the `$query` variable. You may modify it to include/exclude columns, change filters or adapt for other ConfigMgr environments.
- Output paths: modify `$HTMLFileSavePath` and `$CSVFileSavePath` in the script or override by editing the script or using a wrapper that updates those variables prior to invocation.
- Email formatting: the script uses a dedicated `$Body` HTML snippet. Edit as needed for branding or additional information.

---

## Security Considerations

- Do not store plaintext credentials in a repository or unprotected configuration file.
- Secure the config XML location (NTFS ACLs) so only the service account and administrators can read it.
- Consider using certificate-based SMTP, service accounts and restricted database accounts.

---

## Author & License

- Original author: Christian Damberg  
- Last updated: 2025-05-19 (updated by GitHub Copilot)

No license file is included in the repository. If you intend to redistribute or use this script broadly, add an appropriate LICENSE file (for example MIT or another open source / corporate license) and ensure compliance with your organizational policies.

---

If you want, I can:
- Add a sample `ScriptConfigDevInvent.xml` file to the repository,
- Add a simple wrapper to safely load credentials from Windows Credential Manager,
- Or create a small troubleshooting checklist to ship alongside this README.
