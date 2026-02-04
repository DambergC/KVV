# All with client check PASSED and ACTIVE
$results | Where-Object { $_.ClientCheckStatus -eq 'Passed' -and $_.ClientActivityStatus -eq 'Active' }

# All with client check FAILED and INACTIVE
$results | Where-Object { $_.ClientCheckStatus -eq 'Failed' -and $_.ClientActivityStatus -eq 'Inactive' }