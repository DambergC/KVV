# Send-DPstatus.ps1 (DP Status / Maintenance Mode Automation)

## Overview

`Send-DPstatus.ps1` is intended to run as a **scheduled task** on a Microsoft Endpoint Configuration Manager (MECM/SCCM) site server.

It performs a health check of all **Distribution Points (DPs)** by pinging each DP server and then:

- **If a DP is offline (no ping response)**:
  - Enables **Maintenance Mode**
  - Moves the DP from the **Production DP group** to the **Maintenance DP group**
  - Logs the action
  - Includes the DP in an HTML status report as **Offline**
- **If a DP is online**:
  - If it is currently in **Maintenance Mode**, it:
    - Disables **Maintenance Mode**
    - Moves the DP back to the **Production DP group**
    - Logs the action
    - Includes the DP in the report as **Restored**
  - If it is **not** in Maintenance Mode, it is left unchanged and is not added to the report

Finally, the script generates an **HTML email report** (via `ConvertTo-Html`) and sends it using the **Send-MailKitMessage** PowerShell module.

---

## What the script does (high-level flow)

1. Loads required PowerShell modules:
   - `ConfigurationManager` (MECM)
   - `Send-MailKitMessage`
   - `PSWriteHTML` *(imported, but not actively used in the current script body)*

2. Connects to the MECM site by:
   - Detecting **Site Code** via WMI (`Get-CMSiteCode`)
   - Switching to the MECM PSDrive (e.g., `ABC:`)

3. Retrieves all Distribution Points:
   - `Get-CMDistributionPoint -SiteCode <sitecode>`

4. For each DP:
   - Normalizes DP name from `NetworkOSPath`
   - Pings it (`Test-Connection`)
   - Builds two lists:
     - `$SucceededDPs` (online)
     - `$FailedDps` (offline)

5. For offline DPs:
   - Checks MaintenanceMode using `Get-CMDistributionPointInfo`
   - Sets Maintenance Mode **ON**
   - Moves DP to the Maintenance group (`Add-CMDistributionPointToGroup`, `Remove-CMDistributionPointFromGroup`)
   - Adds DP to report result list with Status `Offline`

6. For online DPs:
   - If they are in Maintenance Mode:
     - Sets Maintenance Mode **OFF**
     - Moves DP back to Production group
     - Adds DP to report result list with Status `Restored`
   - Else: skips

7. Builds an HTML email body:
   - Header + embedded logo (base64 image)
   - Totals: total DPs, ping success, ping failed
   - Table produced by `ConvertTo-Html`

8. Sends the mail using `Send-MailKitMessage`.

9. Writes logging to a logfile (path configured near top of script).

---

## Functions

### 1) `Write-Log`

**Purpose:**  
Writes a timestamped entry to the logfile defined in `$LogFile`.

**Input:**
- `-LogString` (string): The message to log

**Behavior:**
- Prepends timestamp in format `yyyy/MM/dd HH:mm:ss`
- Appends line to log file using `Add-Content`

**Example:**
```powershell
Write-Log -LogString "DP01 not online"
```

---

### 2) `Get-CMSiteCode`

**Purpose:**  
Finds the MECM **Site Code** by querying WMI class `SMS_ProviderLocation` in namespace `root\SMS` on the configured site server.

**Output:**
- Returns a string SiteCode, e.g. `ABC`

**Dependencies:**
- `$SiteServer` must be resolvable and accessible via WMI
- Permissions to query WMI on site server

**Example:**
```powershell
$sitecode = Get-CMSiteCode
Set-Location "$sitecode`:"
```

---

## Key configuration values to update

These are set near the top of the script and must match your environment.

### MECM / Site server settings
- `$siteserver`
- `$ProviderMachineName` *(declared but not used in current script logic)*
- `$DPMaintGroup` (name of DP group for maintenance)
- `$DPProdGroup` (name of DP group for production)

### Mail settings
- `$MailFrom`
- `$MailTo1..$MailTo5`
- `$MailSMTP`
- `$MailPortnumber`
- `$MailCustomer` *(declared but not used in current script logic)*

### Log settings
- `$Logfile`  
  Example: `G:\Scripts\Logfiles\DPLogfile.log`

---

## Prerequisites

### 1) MECM Console / ConfigurationManager module
The script imports:

- `ConfigurationManager.psd1` from:
  - `$Env:SMS_ADMIN_UI_PATH`

So it must run on a machine with the MECM console installed and the environment variable available.

### 2) PowerShell modules
- **Send-MailKitMessage**
  - https://github.com/austineric/Send-MailKitMessage
- **PSWriteHTML**
  - https://github.com/EvotecIT/PSWriteHTML *(imported but not required unless you expand reporting)*

Install (example):
```powershell
Install-Module Send-MailKitMessage -Scope AllUsers
Install-Module PSWriteHTML -Scope AllUsers
```

### 3) Permissions
The account running the scheduled task needs:
- Rights to run MECM cmdlets:
  - `Get-CMDistributionPoint`
  - `Get-CMDistributionPointInfo`
  - `Set-CMDistributionPoint`
  - `Add-CMDistributionPointToGroup`
  - `Remove-CMDistributionPointFromGroup`
- File permissions to write the logfile path
- Network access to ping DPs
- SMTP relay permission to send via `$MailSMTP`

---

## Output

### 1) Log file
A logfile is appended to at:
- `$Logfile` (configured path)

### 2) Email report
An HTML email is sent with:
- Subject: `Kontroll av Distributionspunkter - <yyyy-MM-dd>`
- HTML body containing:
  - Totals for online/offline
  - Table of handled DPs (Offline/Restored)

---

## Notes / Known issues (recommended fixes)

These are behaviors to be aware of in the current script:

1. **Assignment used instead of comparison**
   In the offline section, this line uses `=` instead of `-eq`:
   ```powershell
   If($DPStatus.MaintenanceMode = '0'){
   ```
   That assigns instead of compares and will produce unintended behavior. It should be:
   ```powershell
   If ($DPStatus.MaintenanceMode -eq 0) {
   ```

2. **`$limit` and `dateposted` are referenced but not defined**
   The HTML body creation includes:
   ```powershell
   $result | Sort-Object dateposted -Descending | where-object { $_.dateposted -ge $limit }
   ```
   The objects created do not include `dateposted`, and `$limit` is never set. In practice this can cause errors or empty output.
   Recommended: remove these filters or add a timestamp property when creating `$object`.

3. **PSWriteHTML is imported but not used**
   The report is produced via `ConvertTo-Html` and custom CSS.

4. **Credential block is defined but commented out in Parameters**
   A PSCredential object is created but the `"Credential"` parameter is commented out. If your SMTP requires auth, you must enable it.

---

## Suggested scheduling

Run as a scheduled task on the MECM site server, for example:
- Every 15/30/60 minutes depending on operational requirements

Run with:
- A service account with MECM admin rights (or delegated DP admin rights)
- “Run whether user is logged on or not”
- “Run with highest privileges”

---

## Change history

- Initial version: DP ping check + maintenance automation + email report.

---

## Maintainer

- Script owner: `DambergC`
- Organization/customer: Kriminalvården (as configured in script variables)
