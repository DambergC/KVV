# PowerShell script to send Windows Update report

# Required module
Import-Module PSWriteHTML

# Variables
$UpdatesFoundHtml = New-HTML -Title 'Windows Update Report' -Css 'path/to/css' {
    New-HTMLSection -Header 'Updates Found' {
        New-HTMLTable -Data $Updates | Add-HTMLTableHeader -Header @('Update', 'Status', 'Installed')
    }
}

# Preserving existing content
# Month/Year title, intro text, created-from host...

# Email sending logic
$EmailParams = @{...} # existing parameters
$HTMLBody = $UpdatesFoundHtml
Send-MailMessage @EmailParams
