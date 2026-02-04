# Define the task name
$TaskName = "BackupRestoreOutlookSignatures_UserContext"

# Try to remove the scheduled task
try {
    # Check if the scheduled task exists
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        # Remove the task
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-Host "Scheduled Task '$TaskName' removed successfully."
    } else {
        Write-Host "Scheduled Task '$TaskName' does not exist. Nothing to remove."
    }
} catch {
    Write-Error "Failed to remove Scheduled Task '$TaskName': $_"
}
