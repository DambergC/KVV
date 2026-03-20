# Clean IIS Logs (PowerShell)

This folder contains a PowerShell script to clean up IIS log files older than a retention period.

## Script

- **Script name:** `Clean-IISLogs.v2.ps1`
- **Purpose:** Deletes IIS log files older than `KeepDays` from IIS log directories (discovered from IIS config when possible, plus sensible defaults).
- **Automation-friendly:** Non-interactive by default, supports `-WhatIf`/`-Confirm`, optional log file output, and returns structured objects (unless disabled).

## Features

- Keeps the most recent **N days** of logs (default: **30**).
- Discovers IIS site log directories via the **WebAdministration** module (`IIS:\Sites`) when available.
- Always includes the default IIS log directory:
  - `%SystemDrive%\inetpub\logs\LogFiles`
- Optionally includes Failed Request Tracing logs:
  - `%SystemDrive%\inetpub\logs\FailedReqLogFiles`
- Deletes by extension (default `log`, configurable).
- Supports PowerShell safety controls:
  - `-WhatIf` to preview
  - `-Confirm` to prompt
- Optional text logging via `-LogPath`.
- Emits structured output objects per matched/deleted/failed item (can be disabled with `-NoPassThru`).
- Safety rail: skips “dangerous” directories (like `C:\` drive root) unless `-Force` is provided.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+ should work.
- For IIS directory auto-discovery:
  - IIS server with **WebAdministration** module available.
- Administrator privileges are recommended (permissions may prevent deleting in IIS log locations).

## Parameters

### `-KeepDays` (int)
Days of logs to keep. Files older than `(Get-Date).AddDays(-KeepDays)` are targeted.

- Default: `30`
- Range: `1` to `3650`

### `-ExtraPaths` (string[])
Additional directories to scan (optional).

Example:
- `-ExtraPaths "D:\IISLogs","E:\ArchivedLogs"`

### `-IncludeFailedReqLogs` (switch)
Include Failed Request Tracing logs directory.

### `-Extensions` (string[])
File extensions to target, **without the dot**.

- Default: `@('log')`
- Example: `-Extensions log,zip`

### `-NoRecurse` (switch)
Do not recurse into subfolders. (By default the script recurses, because IIS logs are commonly in subfolders such as `W3SVC*`.)

### `-LogPath` (string)
Append a timestamped log file of script activity.

Example:
- `-LogPath "C:\Logs\iis-cleanup.log"`

### `-Force` (switch)
Allow scanning/deleting in paths considered dangerous (like a drive root). Use with care.

### `-NoPassThru` (switch)
Do not emit structured output objects. Useful for “quiet” runs.

## Outputs

By default, the script outputs objects like:

- `Action` (`Matched`, `Deleted`, `Skipped`, `Failed`, `DirectorySkipped`)
- `FullName`
- `Directory`
- `Extension`
- `LastWriteTime`
- `Length`
- `Message`
- `Error`

This makes it easy to export results:

```powershell
.\Clean-IISLogs.v2.ps1 -KeepDays 30 -Confirm:$false |
  Export-Csv -NoTypeInformation -Path .\iis-cleanup-results.csv
```

## Usage Examples

### Preview what would be deleted (recommended first run)

```powershell
.\Clean-IISLogs.v2.ps1 -KeepDays 30 -WhatIf -Verbose
```

### Delete old `.log` files without prompting

```powershell
.\Clean-IISLogs.v2.ps1 -KeepDays 30 -Confirm:$false
```

### Delete `.log` and `.zip` archives, include Failed Request logs, write a log file

```powershell
.\Clean-IISLogs.v2.ps1 `
  -KeepDays 14 `
  -Extensions log,zip `
  -IncludeFailedReqLogs `
  -LogPath "C:\Logs\iis-cleanup.log" `
  -Verbose `
  -Confirm:$false
```

### Include additional custom log paths

```powershell
.\Clean-IISLogs.v2.ps1 -KeepDays 60 -ExtraPaths "D:\IISLogs" -Confirm:$false
```

## Scheduling (Task Scheduler)

Typical arguments for a daily scheduled cleanup:

```powershell
-NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\Clean-IISLogs.v2.ps1" -KeepDays 30 -Confirm:$false -LogPath "C:\Logs\iis-cleanup.log"
```

Recommendation:
- Run with an account that has permissions to the IIS log directories.
- Use `-WhatIf` for a test run before enabling deletions.

## Notes / Safety

- Always validate the discovered directories printed in verbose output before enabling deletion.
- Start with `-WhatIf`.
- Use `-Force` only if you explicitly need to target paths that might be considered risky.

## License

If this repository has a root license, it applies. Otherwise, add a license file if you plan to share/distribute this script.
