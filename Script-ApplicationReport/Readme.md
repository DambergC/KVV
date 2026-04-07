# Script-ApplicationReport

Sends **per-application software installation summaries** from **Microsoft SCCM / ConfigMgr** as HTML reports and emails.

This solution is driven by an XML configuration file (`Send-ApplicationReport.xml`) and produces one report per configured application.

## What's new / key features (script v2.6)

- **Per-application scheduling** via `SendDays="Mon,Tue,..."`.
- **Deployment target reporting** (collection-based) per application:
  - Supports **multiple target collections** via `<TargetCollections><CollectionId>...</CollectionId></TargetCollections>`.
  - Supports **Device** and **User** membership via `TargetType="Device|User"`.
  - **NEW:** Includes **CollectionName** for each configured `CollectionId` (looked up from `v_Collection`).
- **Per-recipient attachment control** via `attach="true|false"` combined with the `-AttachHTML` switch.
- **Dry run mode** (`-DryRun`) to preview HTML without sending emails.
- **Mail-only mode** (`-MailOnly`) to send HTML in the email body without writing files.

---

## Files

- `Send-ApplicationReport.ps1`
  - Reads configuration from the XML
  - Queries the ConfigMgr database using **dbatools** (`Invoke-DbaQuery`)
  - Builds HTML using **PSWriteHTML**
  - Sends email using **Send-MailKitMessage**

- `Send-ApplicationReport.xml` (repo currently shows `Send-ApplicationReport.XML`)
  - Stores SQL server, SMTP settings, output folder, and per-application configuration

> Note on casing: Windows is case-insensitive, but Git is not. The script loads `Send-ApplicationReport.xml` while the repo currently contains `Send-ApplicationReport.XML`. Rename the file or update the script so the names match.

---

## Prerequisites

### PowerShell modules (required)

The script expects these modules to be installed:

- `dbatools` (SQL querying via `Invoke-DbaQuery`)
- `PSWriteHTML` (HTML generation)
- `Send-MailKitMessage` (SMTP mail sending using MailKit/MimeKit)

Install (example):

```powershell
Install-Module dbatools -Scope CurrentUser
Install-Module PSWriteHTML -Scope CurrentUser
Install-Module Send-MailKitMessage -Scope CurrentUser
```

### Access required

- Permissions to query the ConfigMgr SQL database (script DB name is currently hard-coded as `CM_KV1`).
- Ability to send mail via the configured SMTP server.
- Write access to `<HTMLfilePath>` if you want HTML files saved (not used when running with `-MailOnly`).

---

## Configuration: `Send-ApplicationReport.xml`

### Global settings

```xml
<Configuration>
  <HTMLfilePath>G:\Scripts\OutFiles\</HTMLfilePath>
  <SQLServer>vntsql0299.kvv.se</SQLServer>
  <Mailfrom>no-reply@kvv.se</Mailfrom>
  <MailSMTP>smtp.kvv.se</MailSMTP>
  <MailPort>25</MailPort>
  <MailCustomer>Kriminalvarden - IT</MailCustomer>
  ...
</Configuration>
```

**Fields:**

- `HTMLfilePath` — output folder for saved HTML (only when **not** running `-MailOnly`).
- `SQLServer` — SQL Server hosting the ConfigMgr database.
- `Mailfrom` — sender address.
- `MailSMTP` / `MailPort` — SMTP host and port.
- `MailCustomer` — customer label shown in the HTML header.

### Per-application settings

Each application is defined under `<Applications>`:

```xml
<Application Name="Google Chrome" SendDays="Fri" TargetType="Device">
  <ProductFilter>%Google Chrome%</ProductFilter>
  <TargetCollections>
    <CollectionId>KV10039C</CollectionId>
    <CollectionId>KV10034F</CollectionId>
  </TargetCollections>
  <Recipients>
    <Recipient email="user@domain" attach="false" />
  </Recipients>
</Application>
```

**Attributes / elements:**

- `Name` (required) — used in:
  - report title
  - email subject
  - output file name (sanitized)

- `SendDays` (optional) — comma-separated list of allowed send days.
  - Supported: `Mon,Tue,Wed,Thu,Fri,Sat,Sun`
  - If set and **today is not included**, the application is skipped.

- `TargetType` (optional; default `Device`) — determines which membership view is used:
  - `Device` → `v_FullCollectionMembership`
  - `User` → `v_FullCollectionMembership_User`

- `ProductFilter` (required) — used in SQL as:
  - `s.ProductName0 LIKE '<ProductFilter>'`

- `TargetCollections` (optional) — one or more collection IDs to report deployment targets.
  - Output includes `CollectionId`, `CollectionName`, and `TargetCount`, plus a **UniqueTargets** number across the listed collections.

- `Recipients` (required for sending mail)
  - `email` (required)
  - `attach` (optional) — `true|false` (case-insensitive)

---

## Running the script

Run from the folder containing both files:

```powershell
.\Send-ApplicationReport.ps1
```

### Parameters / switches

- `-DryRun`
  - Generates HTML and opens it in the default browser.
  - **Does not send emails.**

- `-MailOnly`
  - Sends HTML in the email body only.
  - **Does not write** HTML files to disk.

- `-AttachHTML`
  - Enables attachment support.
  - A recipient only gets an attachment when:
    - `-AttachHTML` is provided, **and**
    - the recipient has `attach="true"` in XML, **and**
    - an HTML file was actually written (i.e., not `-MailOnly`).

### Examples

```powershell
# Default: generate reports, save HTML, send emails
.\Send-ApplicationReport.ps1

# Preview HTML only (no emails)
.\Send-ApplicationReport.ps1 -DryRun

# Send emails but don't save HTML files
.\Send-ApplicationReport.ps1 -MailOnly

# Allow attachments (still controlled by per-recipient attach flag)
.\Send-ApplicationReport.ps1 -AttachHTML
```

---

## Output

### HTML report

For each application, the HTML includes:

- Generated timestamp
- Customer name
- Filter used
- Total installations (sum of `InstallCount`)
- A results table (Publisher, ProductName, ProductVersion, InstallCount) or a "No matching installations" message
- **Deployment targets per collection** table (under the software table), including **CollectionName**

### Email

- **Subject:** `"<AppName> Software Installation Summary - <yyyy-MM-dd HH:mm>"`
  - When targets are configured and successfully queried, it also includes: `(Targets: <UniqueTargets> Devices|Users)`
- **Body:** HTML report
- **Attachment (optional):** saved HTML file (controlled by `-AttachHTML` + per-recipient `attach="true"`)

---

## Troubleshooting

- **Required module missing**
  - Install: `dbatools`, `PSWriteHTML`, `Send-MailKitMessage`

- **Application is skipped unexpectedly**
  - Check `SendDays` values are valid and include today's day (Mon..Sun).

- **No data in report**
  - Verify `ProductFilter` matches `v_GS_INSTALLED_SOFTWARE.ProductName0` values.

- **Deployment target counts missing**
  - Verify `TargetType` and `CollectionId` values.
  - Ensure the ConfigMgr views exist and permissions allow reading them.

- **Attachments not included**
  - Ensure `-AttachHTML` is used.
  - Ensure recipient has `attach="true"`.
  - Ensure HTML files are being written (not `-MailOnly`, and `HTMLfilePath` is valid and writable).