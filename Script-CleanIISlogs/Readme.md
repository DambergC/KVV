# Clean-IISLogs.ps1

Deletes IIS log files older than a configured retention period and appends actions to a log file.

## What it does
- Recursively searches for `*.log` under the configured log root.
- Selects files with `LastWriteTime` older than the cutoff date (`Get-Date).AddDays($maxDaysToKeep)`).
- Logs each file that will be deleted.
- Deletes the file.
- If nothing is found, writes a "No items to be deleted" entry.

## Configuration
Edit these variables at the top of `Clean-IISLogs.ps1`:

- `$LogPath` – Root folder where IIS logs are stored (default: `C:\inetpub\logs`).
- `$maxDaystoKeep` – Negative number of days to keep logs. Example: `-12` keeps the last 12 days and deletes anything older.
- `$outputPath` – Path to the cleanup log file.

## Important notes
### 1) `Add-Content` requires the output directory to exist
`Add-Content` can create the log file, but it cannot create missing parent folders. Ensure the directory in `$outputPath` exists.

Suggested snippet to add near the top of the script:

```powershell
$logDir = Split-Path -Path $outputPath -Parent
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}
