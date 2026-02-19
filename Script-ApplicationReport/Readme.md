# Script-ApplicationReport

This folder contains a PowerShell script and an XML configuration file used to generate **per-application software installation summaries** from **Microsoft SCCM / ConfigMgr** inventory data and distribute them as **HTML email reports**.

## Files

- `Send-ApplicationReport.ps1`  
  Reads configuration from the XML, queries the ConfigMgr database, generates an HTML report, and emails it to recipients.

- `Send-ApplicationReport.xml`  
  Configuration file containing SQL connection details, SMTP/mail settings, output path for HTML files, and per-application recipient lists (including per-recipient attachment control).

> Note: Windows is case-insensitive, but the repo currently shows `Send-ApplicationReport.XML`. The script loads `Send-ApplicationReport.xml`. Ensure the filename matches what the script expects, or update the script accordingly.

---

## What the script does (high level)

For each configured `<Application>` in the XML:

1. Builds a SQL query using the `<ProductFilter>` (used in `WHERE ProductName0 LIKE '<ProductFilter>'`).
2. Runs the query against the ConfigMgr database (`CM_KV1`) using **dbatools** (`Invoke-DbaQuery`).
3. Builds an HTML report using **PSWriteHTML**:
   - A header section with generated timestamp, customer name, filter used, and total installation count.
   - A table of results (Publisher, ProductName, ProductVersion, InstallCount) or a “No matching installations found” message.
4. Optionally writes the HTML report to disk (unless `-MailOnly` is used).
5. Sends **one email per recipient** using **Send-MailKitMessage**:
   - Always includes the HTML report in the email body.
   - Optionally attaches the generated HTML file depending on:
     - the global `-AttachHTML` switch, **and**
     - the recipient’s `attach="true"` attribute in the XML, **and**
     - a saved HTML file exists.

---

## Prerequisites

### PowerShell modules (required)
The script checks for these modules and exits if any are missing:

- `dbatools` (used for `Invoke-DbaQuery`)
- `PSWriteHTML` (used to generate the HTML)
- `Send-MailKitMessage` (used to send email via SMTP using MailKit/MimeKit types)

Install (example, run in an elevated PowerShell session if required by policy):

```powershell
Install-Module dbatools -Scope CurrentUser
Install-Module PSWriteHTML -Scope CurrentUser
Install-Module Send-MailKitMessage -Scope CurrentUser
```

### Access required
- Network access + permissions to query the ConfigMgr SQL database (default DB name in script: `CM_KV1`).
- Ability to send mail via the configured SMTP server.
- Write access to `<HTMLfilePath>` if you want the script to save HTML files.

---

## Configuration: `Send-ApplicationReport.xml`

### Global settings

```xml
<Configuration>
  <HTMLfilePath>G:\Scripts\OutFiles\</HTMLfilePath>
  <SQLServer>ServerName</SQLServer>
  <Mailfrom>no-reply@shelby.org</Mailfrom>
  <MailSMTP>smtp.shelby.org</MailSMTP>
  <MailPort>25</MailPort>
  <MailCustomer>Shelby Company Limited</MailCustomer>
  ...
</Configuration>
```

**Fields:**

- `HTMLfilePath`  
  Folder where HTML report files are saved (only used when **not** running with `-MailOnly`).

- `SQLServer`  
  SQL Server name / instance hosting the ConfigMgr DB.

- `Mailfrom`  
  Sender email address.

- `MailSMTP`  
  SMTP host.

- `MailPort`  
  SMTP port (cast to integer in script).

- `MailCustomer`  
  Customer label displayed in the HTML header.

### Per-application settings

Applications are defined under:

```xml
<Applications>
  <Application Name="Google Chrome">
    <ProductFilter>Google Chrome%</ProductFilter>
    <Recipients>
      <Recipient email="thomas.shelby@shelby.org" attach="False" />
    </Recipients>
  </Application>
</Applications>
```

**Fields / attributes:**

- `<Application Name="...">` (required)  
  Human-friendly application name used in:
  - the HTML title
  - the email subject
  - the output file name (sanitized)

- `<ProductFilter>...</ProductFilter>` (required)  
  Used in SQL as: `ProductName0 LIKE '<ProductFilter>'`  
  Example: `Google Chrome%`

- `<Recipients>` / `<Recipient ... />` (required for sending email)
  - `email="user@domain"` (required)
  - `attach="true|false"` (optional, evaluated case-insensitively by converting to lowercase and comparing to `true`)

**Attachment behavior:**
- If you do **not** run the script with `-AttachHTML`, **no one** gets attachments.
- If you run with `-AttachHTML`, only recipients with `attach="true"` get the HTML file attached (and only if the HTML file was written to disk).

---

## Running the script

Run from the folder containing both files:

```powershell
.\Send-ApplicationReport.ps1
```

### Switches

- `-DryRun`  
  Does not send any email. Generates HTML and opens it in the default browser.
  - If HTML files are written, it opens the saved file.
  - If `-MailOnly` is also used, it creates a temporary HTML file and opens that.

- `-MailOnly`  
  Sends the HTML in the email body only and **does not write** HTML files to disk.

- `-AttachHTML`  
  Enables attachment support. Attachments are still controlled per-recipient via `attach="true"` in the XML.

### Examples

**Generate + send emails (default behavior):**
```powershell
.\Send-ApplicationReport.ps1
```

**Test the generated HTML without sending emails:**
```powershell
.\Send-ApplicationReport.ps1 -DryRun
```

**Send emails but don’t save HTML files:**
```powershell
.\Send-ApplicationReport.ps1 -MailOnly
```

**Allow attachments (only to recipients with attach="true"):**
```powershell
.\Send-ApplicationReport.ps1 -AttachHTML
```

---

## What data is queried (SQL overview)

The script queries ConfigMgr inventory view `v_GS_INSTALLED_SOFTWARE` joined to `System_DATA`, grouped by:

- Publisher
- ProductName
- ProductVersion

And returns an `InstallCount` per group for rows where:

- `ProductName0 LIKE '<ProductFilter from XML>'`

---

## Output

### HTML report
The report includes:
- Report title: `"<ApplicationName> Software Installation Summary"`
- Generated timestamp
- Customer name from XML
- Filter used
- Total installations (sum of `InstallCount`)
- A table with:
  - Publisher
  - ProductName
  - ProductVersion
  - InstallCount

### Email
- **Subject:** `"<ApplicationName> Software Installation Summary - <yyyy-MM-dd HH:mm>"`
- **Body:** HTML report content
- **Attachment (optional):** the saved HTML file (per rules above)

---

## Troubleshooting

- **Missing module error**
  - Install the missing module(s): `dbatools`, `PSWriteHTML`, `Send-MailKitMessage`.

- **No emails sent for an application**
  - Ensure `<Recipients><Recipient ... /></Recipients>` exists and each recipient has an `email="..."`.

- **No data in the report**
  - Verify `<ProductFilter>` matches the values in `v_GS_INSTALLED_SOFTWARE.ProductName0`.
  - Test the generated SQL filter in SQL Server Management Studio.

- **Attachments not included**
  - Ensure you ran with `-AttachHTML`.
  - Ensure recipient has `attach="true"`.
  - Ensure HTML files are being written (i.e., you did **not** use `-MailOnly`, and `HTMLfilePath` is set and writable).

---

## Notes / customization points

- The ConfigMgr database name is currently hard-coded in the script as `CM_KV1`. If your site DB name differs, update the script or make it configurable via XML.
- `HTMLfilePath` is only used when saving reports; email body is always sent as HTML unless using `-DryRun`.
