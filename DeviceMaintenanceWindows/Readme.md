# DeviceMaintenanceWindows

Scripts and helper functions to create, manage, and document **Windows device maintenance windows** (typically for patching/reboots/servicing) in a consistent way.

> This folder is intended to be used from PowerShell. The scripts are written to be composable: you can run a main script end-to-end, or call individual functions from your own automation.

---

## What this does (high level)

Depending on which script(s) in this folder you run, the tooling generally helps you:

- Define a **maintenance window** (start/end time, time zone, recurrence if applicable).
- Apply that window to a **set of devices** (often via a group, a tag, a query, or an input list).
- Optionally produce **output** (logs / CSV / JSON) describing what was changed and when.

If you have multiple environments (Prod/Test) or multiple regions, the functions are meant to make maintenance windows repeatable and auditable.

---

## Prerequisites

### 1) PowerShell version
- **Windows PowerShell 5.1** *or* **PowerShell 7+** is typically fine.
- Recommended: **PowerShell 7+** for better TLS defaults and cross-platform behavior (if you run from a build agent).

Check:
```powershell
$PSVersionTable.PSVersion
```

### 2) Execution policy
If scripts are blocked on Windows, you may need one of the following (choose what matches your security policy):

```powershell
# Current session only
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Or for the current user
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3) Required modules (common patterns)
Because I can’t see the exact module imports in your scripts from this chat alone, here are the **most common** prerequisites for this kind of automation. Check your `.ps1` / `.psm1` files for `Import-Module` lines and install what they require.

Typical examples:

- `Microsoft.Graph.*` (if managing device groups / Intune / Entra via Graph)
- `Az.Accounts` / `Az.Resources` (if targeting Azure resources)
- `PSReadLine` (usually optional)
- Any **internal module** shipped in this repo (e.g., `.\KVV.*.psm1`)

Install examples:
```powershell
# Microsoft Graph (example)
Install-Module Microsoft.Graph -Scope CurrentUser

# Az modules (example)
Install-Module Az -Scope CurrentUser
```

### 4) Authentication / permissions
You will typically need:
- Rights to **read** the target device inventory and groups
- Rights to **update** whatever entity stores your “maintenance window” configuration (device group metadata, config profile, tags, etc.)
- If using Microsoft Graph: the relevant **Graph scopes** (e.g., DeviceManagement*, Group.ReadWrite.All, etc.)

> Tip: run once interactively to sign in, then use a service principal / managed identity for automation if needed.

---

## Folder contents (expected layout)

You will likely see one or more of the following types of files:

- `*.ps1` — runnable scripts (entry points)
- `*.psm1` — function modules (reusable functions)
- `*.psd1` — module manifests
- `*.json` / `*.csv` — configuration or input data

If there is a “main” script, it’s usually the one with a name like:
- `Invoke-*.ps1`
- `New-*.ps1`
- `Set-*.ps1`
- `Start-*.ps1`

---

## Functions / scripts (what they do)

### Common function responsibilities
Even if your exact function names differ, most maintenance-window toolsets include functions in these categories:

1. **Input & validation**
   - Validate date/time formats, end > start, allowed time zones, etc.
   - Validate device identifiers (names, IDs) and that targets exist.

2. **Target selection**
   - Resolve devices from:
     - a group
     - a CSV
     - a query/filter
     - explicit device IDs

3. **Maintenance window calculation**
   - Normalize to a single time zone or UTC.
   - Calculate next occurrence for recurring windows.
   - Clamp/round times (e.g., to 15-minute boundaries) if required by an API/system.

4. **Apply / update**
   - Create or update the maintenance window object/setting.
   - Assign it to devices/groups.
   - Handle idempotency (re-running shouldn’t duplicate settings).

5. **Logging & output**
   - Write summary to console
   - Write detailed logs to a file
   - Export results to CSV/JSON

---

## How to use

### Option A — Run the main script (recommended)
1) Open PowerShell in this folder:
```powershell
cd .\DeviceMaintenanceWindows
```

2) If scripts depend on functions in a module in this folder, import it first (example):
```powershell
Import-Module .\DeviceMaintenanceWindows.psm1 -Force
```

3) Run the script (example patterns):
```powershell
# Example
.\Invoke-DeviceMaintenanceWindow.ps1

# Or with parameters (example)
.\Invoke-DeviceMaintenanceWindow.ps1 -Environment Prod -TimeZone "W. Europe Standard Time"
```

> Use `Get-Help` to discover parameters if comment-based help exists:
```powershell
Get-Help .\Invoke-DeviceMaintenanceWindow.ps1 -Full
```

### Option B — Use functions from your own automation
If the folder provides a module (`.psm1`), import and call functions:

```powershell
Import-Module .\DeviceMaintenanceWindows.psm1 -Force

# Example function calls (names will differ in your repo)
$window = New-MaintenanceWindow -Start "2026-02-10 22:00" -End "2026-02-11 02:00" -TimeZone "UTC"
Set-DeviceMaintenanceWindow -TargetGroupId "<group-guid>" -MaintenanceWindow $window
```

---

## Typical inputs

### Date/time
Prefer explicit, unambiguous values:
- ISO-like: `2026-02-10 22:00`
- Or UTC with `Z` if your functions accept it: `2026-02-10T22:00:00Z`

### Device lists (CSV)
A common pattern is a CSV with one of these columns:
- `DeviceName`
- `DeviceId`
- `SerialNumber`

Example CSV:
```csv
DeviceName
PC-001
PC-002
PC-003
```

---

## Outputs / logs

Look for any of these patterns:
- A `.\logs\` folder
- A transcript file started with `Start-Transcript`
- Exports like `Export-Csv` / `ConvertTo-Json`

If nothing else exists, you can wrap execution with a transcript:
```powershell
Start-Transcript -Path .\maintenancewindow-transcript.txt -Force
.\Invoke-DeviceMaintenanceWindow.ps1
Stop-Transcript
```

---

## Troubleshooting

### “Running scripts is disabled on this system”
Use an appropriate execution policy (see Prerequisites).

### Authentication failures / forbidden
- Ensure you are signed in to the correct tenant/subscription.
- Verify required roles/scopes.
- If using Graph: confirm consent was granted for the scopes your script requests.

### Time zone confusion
- Confirm what the script expects (local time vs UTC).
- If your org spans regions, standardize on UTC and convert for display only.

### Idempotency / duplicates
If re-runs create duplicates, look for flags like:
- `-WhatIf`
- `-Confirm:$false`
- `-Force`
- `-UpdateExisting`

---

## Notes / customization

- If you have multiple window templates (e.g., “Pilot”, “Broad”, “Servers”), consider keeping them as JSON configs and selecting by name.
- For CI/CD, prefer non-interactive auth and avoid prompting (set `-Confirm:$false` and provide all parameters explicitly).

---

## Contributing

- Keep functions small and composable.
- Add comment-based help to scripts/functions:
  - `.SYNOPSIS`
  - `.DESCRIPTION`
  - `.PARAMETER`
  - `.EXAMPLE`
- Include a dry-run path (`-WhatIf`) for safety when changing assignments at scale.
