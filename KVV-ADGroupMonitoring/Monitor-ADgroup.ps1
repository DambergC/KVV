param (
    [string]$GroupName,
    [switch]$Test
)

# Ensure PSWriteHTML module is installed for all users
if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
    Install-Module -Name PSWriteHTML -Scope AllUsers -Force
}

# Import the module
Import-Module PSWriteHTML

# Ensure Send-MailKitMessage module is installed for all users
if (-not (Get-Module -ListAvailable -Name Send-MailKitMessage)) {
    Install-Module -Name Send-MailKitMessage -Scope AllUsers -Force
}

# Import the module
Import-Module Send-MailKitMessage

# Get the script's directory
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Define folder paths
$InventoryFolder = Join-Path -Path $ScriptDir -ChildPath "MonitorADgroupInventory"
$ChangeHistoryFolder = Join-Path -Path $ScriptDir -ChildPath "MonitorADgroupChange"
$HtmlReportPath = Join-Path -Path $ScriptDir -ChildPath "MonitorADgroupReports\$GroupName.html"
$HtmlEmailPath = Join-Path -Path $ScriptDir -ChildPath "MonitorADgroupReports\$GroupName-email.html"

# Create folders if they don't exist
if (-not (Test-Path $InventoryFolder)) {
    New-Item -Path $InventoryFolder -ItemType Directory
}

if (-not (Test-Path $ChangeHistoryFolder)) {
    New-Item -Path $ChangeHistoryFolder -ItemType Directory
}

if (-not (Test-Path (Split-Path $HtmlReportPath))) {
    New-Item -Path (Split-Path $HtmlReportPath) -ItemType Directory
}

# Function to get current group members
function Get-CurrentGroupMembers {
    $GroupMembers = Get-ADGroupMember -Identity $GroupName
    $GroupMembers | ForEach-Object {
        $User = Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, DistinguishedName
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $_.SamAccountName
            DistinguishedName = $User.DistinguishedName
        }
    }
}

# Function to save group members to CSV
function Save-GroupMembersToCsv {
    param (
        [string]$FilePath,
        [array]$Members,
        [switch]$Append
    )
    if ($Append) {
        $Members | Export-Csv -Path $FilePath -NoTypeInformation -Append
    } else {
        $Members | Export-Csv -Path $FilePath -NoTypeInformation
    }
}

# Function to generate HTML report
function Generate-HTMLReport {
    param (
        [array]$Changes,
        [string]$FilePath,
        [array]$History
    )
    $HtmlContent = New-HTML {
        # Top header
        New-HTMLSection -Title "Status - Membership - $GroupName - $(Get-Date -Format 'yyyy-MM-dd')" {}

        # Changes from last report
        New-HTMLSection -Title "Changes from Last Report" {
            if ($Changes) {
                New-HTMLTable -DataTable $Changes -Title "AD Group Changes Report for $GroupName"
            } else {
                "No changes detected."
            }
        }

        # Last 10 changes
        New-HTMLSection -Title "Last 10 Changes" {
            if ($History) {
                New-HTMLTable -DataTable $History -Title "Last 10 Changes"
            } else {
                "No history changes found."
            }
        }
    }
    Save-HTML -HTML $HtmlContent -FilePath $FilePath
}

# Function to send email using Send-MailKitMessage
function Send-MailKitMessage {
    param (
        [string]$HtmlFilePath,
        [string]$To,
        [string]$From,
        [string]$Subject,
        [string]$SmtpServer,
        [int]$SmtpPort,
        [string]$SmtpUser,
        [string]$SmtpPassword
    )
    $HtmlBody = Get-Content -Path $HtmlFilePath -Raw
    $Message = New-Object MimeKit.MimeMessage
    $Message.From.Add($From)
    $Message.To.Add($To)
    $Message.Subject = $Subject
    $Message.Body = New-Object MimeKit.TextPart('html') -Property @{ Text = $HtmlBody }

    $Client = New-Object MailKit.Net.Smtp.SmtpClient
    $Client.Connect($SmtpServer, $SmtpPort, $true)
    $Client.Authenticate($SmtpUser, $SmtpPassword)
    $Client.Send($Message)
    $Client.Disconnect($true)
    $Client.Dispose()
}

# Function to compare current members with previous CSV
function Compare-GroupMembers {
    $CsvFilePath = Join-Path -Path $InventoryFolder -ChildPath "$GroupName.csv"
    $CurrentMembers = Get-CurrentGroupMembers

    if (Test-Path $CsvFilePath) {
        $PreviousMembers = Import-Csv -Path $CsvFilePath

        $AddedMembers = Compare-Object -ReferenceObject $PreviousMembers -DifferenceObject $CurrentMembers -Property SamAccountName | Where-Object { $_.SideIndicator -eq "=>" }
        $RemovedMembers = Compare-Object -ReferenceObject $PreviousMembers -DifferenceObject $CurrentMembers -Property SamAccountName | Where-Object { $_.SideIndicator -eq "<=" }

        if ($AddedMembers -or $RemovedMembers) {
            Write-Host "Changes detected in group members."

            $Changes = @()

            if ($AddedMembers) {
                Write-Host "Members added:"
                $AddedMembers | ForEach-Object {
                    $User = Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, DistinguishedName
                    Write-Host "$($_.SamAccountName) added"
                    $Changes += [PSCustomObject]@{
                        DateTime         = Get-Date
                        State            = "Added"
                        DisplayName      = $User.DisplayName
                        SamAccountName   = $_.SamAccountName
                        DistinguishedName = $User.DistinguishedName
                    }
                }
            }

            if ($RemovedMembers) {
                Write-Host "Members removed:"
                $RemovedMembers | ForEach-Object {
                    $User = Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, DistinguishedName
                    Write-Host "$($_.SamAccountName) removed"
                    $Changes += [PSCustomObject]@{
                        DateTime         = Get-Date
                        State            = "Removed"
                        DisplayName      = $User.DisplayName
                        SamAccountName   = $_.SamAccountName
                        DistinguishedName = $User.DistinguishedName
                    }
                }
            }

            # Save changes to ChangeHistory folder with group name, appending to the file
            $ChangeHistoryFilePath = Join-Path -Path $ChangeHistoryFolder -ChildPath "$GroupName.csv"
            Save-GroupMembersToCsv -FilePath $ChangeHistoryFilePath -Members $Changes -Append

            # Get the last 10 changes and reverse the order
            $HistoryChanges = Import-Csv -Path $ChangeHistoryFilePath | Select-Object -Last 10 | Sort-Object DateTime -Descending

            # Generate HTML report
            Generate-HTMLReport -Changes $Changes -FilePath $HtmlReportPath -History $HistoryChanges

            # Generate HTML email file if Test parameter is specified
            if ($Test) {
                Copy-Item -Path $HtmlReportPath -Destination $HtmlEmailPath
                Write-Host "HTML email file generated at $HtmlEmailPath"
            } else {
                # Send email
                Send-MailKitMessage -HtmlFilePath $HtmlReportPath -To "recipient@example.com" -From "sender@example.com" -Subject "AD Group Changes Report for $GroupName" -SmtpServer "smtp.example.com" -SmtpPort 587 -SmtpUser "smtpuser" -SmtpPassword "smtppassword"
            }
        } else {
            Write-Host "No changes detected in group members."
        }

        # Overwrite the inventory CSV with the current group members
        Save-GroupMembersToCsv -FilePath $CsvFilePath -Members $CurrentMembers
    } else {
        Write-Host "Inventory CSV file not found. Saving current members to CSV."
        Save-GroupMembersToCsv -FilePath $CsvFilePath -Members $CurrentMembers
    }
}

# Run the comparison
Compare-GroupMembers