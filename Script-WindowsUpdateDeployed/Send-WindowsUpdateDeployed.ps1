# Updated script content with new HTML report generation

# Existing content up to the relevant section

# Updating HTML report generation
# Install-Module -Name PowerShellUniversal

# Define parameters
$titleText = "$MailCustomer - Windows Updates $monthname $year"

# Start HTML Document
$html = New-HTML -TitleText $titleText 

# Add embedded image (base64 example)
$image = 'data:image/png;base64,....'
$html += '<img src="' + $image + '" />'

# Adding styles and classes
$html += '<style>body { font-family: Arial; } .small { font-size: 10px; } .summary { font-weight: bold; }</style>'

# Add sections for preText
$preText = "Pre-Report Summary Here."
$html += '<div class="summary">' + $preText + '</div>'

# Create updates table
$updatesTable = '<table><tr><th>Date</th><th>Update</th></tr><tr><td>Date1</td><td>Update1</td></tr></table>'
$html += $updatesTable

# Add sections for postText
$postText = "Post-Report Summary Here."
$html += '<div class="summary">' + $postText + '</div>'

# Set HTMLBody
$HTMLBody = [string]$html

# Existing email subject handling
# Send email logic based on $HTMLBody

# Rest of the existing logic
