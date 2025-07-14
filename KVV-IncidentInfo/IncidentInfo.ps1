<#
.SYNOPSIS
    Information Board for notifying local users about ongoing incidents.
.DESCRIPTION
    This script provides functions to display Windows Toast Notifications to inform users about ongoing incidents.
    It can be used by IT administrators to communicate important information to users during service disruptions,
    maintenance, or other incidents. Notifications will remain visible until dismissed by the user.
.NOTES
    Date Created: 2025-06-25
    Last Updated: 2025-06-25 12:49:09
    Author: DambergC
#>

function Show-IncidentNotification {
    [CmdletBinding()]
    Param (
        [string]
        [Parameter(Mandatory=$true)]
        $Title,
 
        [string]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Message,
        
        [string]
        [Parameter(Mandatory=$false)]
        $ApplicationId = "Incident Management System",
        
        [string]
        [Parameter(Mandatory=$false)]
        $ImagePath,
        
        [switch]
        [Parameter(Mandatory=$false)]
        $UrgentNotification,
        
        [string]
        [Parameter(Mandatory=$false)]
        $SupportButtonLabel = "Contact Support",
        
        [string]
        [Parameter(Mandatory=$false)]
        $SupportButtonUrl = "https://support.example.com",
        
        [switch]
        [Parameter(Mandatory=$false)]
        $ShowSupportButton
    )
    
    # Import required libraries for Windows notifications
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    [Windows.System.User, Windows.System, ContentType = WindowsRuntime] > $null
    [Windows.System.UserType, Windows.System, ContentType = WindowsRuntime] > $null
    [Windows.System.UserAuthenticationStatus, Windows.System, ContentType = WindowsRuntime] > $null
    [Windows.Storage.ApplicationData, Windows.Storage, ContentType = WindowsRuntime] > $null
    
    # Clean application ID (remove spaces)
    $AppId = $ApplicationId -replace '\s+', '.'
    
    # Register the AppID in the registry
    try {
        $RegPath = "HKCU:\SOFTWARE\Classes\AppUserModelId\$AppId"
        if (-not (Test-Path -Path $RegPath)) {
            New-Item -Path $RegPath -Force > $null
        }
        Set-ItemProperty -Path $RegPath -Name "DisplayName" -Value $ApplicationId -Type String -Force
        
        # Enable urgent notifications if specified
        if ($UrgentNotification) {
            $NotifPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\$AppId"
            if (-not (Test-Path -Path $NotifPath)) {
                New-Item -Path $NotifPath -Force > $null
            }
            Set-ItemProperty -Path $NotifPath -Name "AllowUrgentNotifications" -Value 1 -Type DWord -Force
        }
    }
    catch {
        Write-Warning "Failed to set registry keys: $_"
    }
    
    # Create toast notifier
    try {
        $ToastNotifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId)
    }
    catch {
        Write-Error "Failed to create notification manager: $_"
        return
    }
    
    # Prepare actions XML with or without support button
    $actionsXml = "<actions><action activationType='system' arguments='dismiss' content='Dismiss'/>"
    
    if ($ShowSupportButton) {
        $actionsXml += "<action activationType='protocol' arguments='$SupportButtonUrl' content='$SupportButtonLabel'/>"
    }
    
    $actionsXml += "</actions>"
    
    # Create XML template for notification - setting scenario to "reminder" to keep it persistent
    # Using reminder scenario to make the toast persistent until user interaction
    $ToastXml = [xml] @"
<toast scenario="reminder">
    <visual>
        <binding template="ToastGeneric">
            <text hint-maxLines="1">$Title</text>
            <text>$Message</text>
            $(if($ImagePath){ "<image placement='appLogoOverride' src='$ImagePath' hint-crop='circle'/>" })
        </binding>
    </visual>
    $actionsXml
    <audio src="ms-winsoundevent:Notification.Default" loop="false"/>
</toast>
"@
    
    # Serialize XML
    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($ToastXml.OuterXml)
    
    # Setup notification properties
    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "Incident"
    $Toast.Group = "IncidentNotifications"
    
    # Removing ExpirationTime to keep the toast until user dismisses it
    # Note: Toast notifications will still be removed if the system is restarted
    
    # Show notification to user
    try {
        $ToastNotifier.Show($Toast)
        Write-Host "Incident notification displayed to user. Will remain until dismissed."
        return $true
    }
    catch {
        Write-Error "Failed to display notification: $_"
        return $false
    }
}

function Update-IncidentStatus {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Ongoing', 'Resolved', 'Scheduled', 'Investigating')]
        [string]$Status,
        
        [Parameter(Mandatory=$true)]
        [string]$IncidentTitle,
        
        [Parameter(Mandatory=$true)]
        [string]$IncidentMessage,
        
        [Parameter(Mandatory=$false)]
        [string]$EstimatedResolutionTime,
        
        [Parameter(Mandatory=$false)]
        [switch]$UrgentNotification,
        
        [Parameter(Mandatory=$false)]
        [string]$SupportButtonLabel = "Contact Support",
        
        [Parameter(Mandatory=$false)]
        [string]$SupportButtonUrl = "https://support.example.com",
        
        [Parameter(Mandatory=$false)]
        [switch]$ShowSupportButton
    )
    
    $currentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Build the message with the current status
    $fullMessage = "$IncidentMessage`n`nStatus: $Status`nLast Updated: $currentDate"
    
    if ($EstimatedResolutionTime) {
        $fullMessage += "`nEstimated Resolution: $EstimatedResolutionTime"
    }
    
    # Create the notification with the support button if specified
    Show-IncidentNotification -Title $IncidentTitle -Message $fullMessage -UrgentNotification:$UrgentNotification `
                              -ShowSupportButton:$ShowSupportButton -SupportButtonLabel $SupportButtonLabel `
                              -SupportButtonUrl $SupportButtonUrl
}

# Function to display a persistent information board with more customizable options
function Show-InformationBoard {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Title,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$SupportContact = "IT Support Desk: x1234",
        
        [Parameter(Mandatory=$false)]
        [string]$SupportButtonLabel = "Get Support",
        
        [Parameter(Mandatory=$false)]
        [string]$SupportButtonUrl = "https://support.example.com",
        
        [Parameter(Mandatory=$false)]
        [string]$IconPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$Critical,
        
        [Parameter(Mandatory=$false)]
        [switch]$HideSupportButton
    )
    
    $scenario = if ($Critical) { "alarm" } else { "reminder" }
    $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Build a more detailed information board
    $detailedMessage = @"
$Message

Additional Information:
- Posted: $currentTime
- $SupportContact
"@
    
    # Prepare actions XML
    $actionsXml = "<actions><action activationType='system' arguments='dismiss' content='Acknowledge'/>"
    
    if (-not $HideSupportButton) {
        $actionsXml += "<action activationType='protocol' arguments='$SupportButtonUrl' content='$SupportButtonLabel'/>"
    }
    
    $actionsXml += "</actions>"
    
    # Create XML template for the information board
    $ToastXml = [xml] @"
<toast scenario="$scenario">
    <visual>
        <binding template="ToastGeneric">
            <text hint-maxLines="1">$Title</text>
            <text>$detailedMessage</text>
            $(if($IconPath){ "<image placement='appLogoOverride' src='$IconPath' hint-crop='circle'/>" })
        </binding>
    </visual>
    $actionsXml
    <audio src="ms-winsoundevent:Notification.$(if($Critical){"Looping.Alarm2"}else{"Default"})" loop="$($Critical.ToString().ToLower())"/>
</toast>
"@
    
    # Use the same AppId for all information boards
    $infoBoardAppId = "Company.InformationBoard"
    
    # Register the AppID
    $RegPath = "HKCU:\SOFTWARE\Classes\AppUserModelId\$infoBoardAppId"
    if (-not (Test-Path -Path $RegPath)) {
        New-Item -Path $RegPath -Force > $null
    }
    Set-ItemProperty -Path $RegPath -Name "DisplayName" -Value "Company Information Board" -Type String -Force
    
    # Create toast notifier
    try {
        # Import required libraries
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
        [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
        
        $ToastNotifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($infoBoardAppId)
        
        # Serialize XML
        $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $SerializedXml.LoadXml($ToastXml.OuterXml)
        
        # Setup notification
        $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
        $Toast.Tag = "InfoBoard"
        $Toast.Group = "CompanyInformationBoard"
        
        # Show notification
        $ToastNotifier.Show($Toast)
        Write-Host "Information Board displayed to user. Will remain until dismissed."
        return $true
    }
    catch {
        Write-Error "Failed to display Information Board: $_"
        return $false
    }
}

# Function to create a custom information popup with multiple custom buttons
function Show-CustomIncidentPopup {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Title,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [hashtable[]]$CustomButtons,
        
        [Parameter(Mandatory=$false)]
        [string]$IconPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$Critical
    )
    
    # Import required libraries
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    
    # Prepare the scenario based on criticality
    $scenario = if ($Critical) { "alarm" } else { "reminder" }
    
    # Build the XML for custom buttons
    $buttonsXml = "<actions>"
    
    # Always add Dismiss button first
    $buttonsXml += "<action activationType='system' arguments='dismiss' content='Dismiss'/>"
    
    # Add custom buttons if provided
    if ($CustomButtons -and $CustomButtons.Count -gt 0) {
        foreach ($button in $CustomButtons) {
            if ($button.ContainsKey('Label') -and $button.ContainsKey('Url')) {
                $buttonsXml += "<action activationType='protocol' arguments='$($button.Url)' content='$($button.Label)'/>"
            }
        }
    }
    
    $buttonsXml += "</actions>"
    
    # Create XML template
    $ToastXml = [xml] @"
<toast scenario="$scenario">
    <visual>
        <binding template="ToastGeneric">
            <text hint-maxLines="1">$Title</text>
            <text>$Message</text>
            $(if($IconPath){ "<image placement='appLogoOverride' src='$IconPath' hint-crop='circle'/>" })
        </binding>
    </visual>
    $buttonsXml
    <audio src="ms-winsoundevent:Notification.$(if($Critical){"Looping.Alarm2"}else{"Default"})" loop="$($Critical.ToString().ToLower())"/>
</toast>
"@
    
    # Use a custom AppId
    $customAppId = "Company.CustomIncidentPopup"
    
    # Register the AppID
    $RegPath = "HKCU:\SOFTWARE\Classes\AppUserModelId\$customAppId"
    if (-not (Test-Path -Path $RegPath)) {
        New-Item -Path $RegPath -Force > $null
    }
    Set-ItemProperty -Path $RegPath -Name "DisplayName" -Value "Company Incident Notification" -Type String -Force
    
    # Create toast notifier and show notification
    try {
        $ToastNotifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($customAppId)
        
        $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $SerializedXml.LoadXml($ToastXml.OuterXml)
        
        $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
        $Toast.Tag = "CustomIncident"
        $Toast.Group = "CompanyIncidents"
        
        $ToastNotifier.Show($Toast)
        return $true
    }
    catch {
        Write-Error "Failed to display custom incident popup: $_"
        return $false
    }
}

# Examples of usage:

# Example 1: Simple notification with support button
# Show-IncidentNotification -Title "Network Maintenance" -Message "Network maintenance in progress. You may experience intermittent connectivity issues." -ShowSupportButton -SupportButtonLabel "Get Help" -SupportButtonUrl "https://helpdesk.example.com"

# Example 2: Information board for a critical incident
# Show-InformationBoard -Title "SECURITY ALERT" -Message "We are currently experiencing a security incident affecting email systems. Please do not open any suspicious emails." -Critical -SupportContact "Security Team: x5555" -SupportButtonLabel "Report Issue" -SupportButtonUrl "https://security.example.com/report"

# Example 3: Custom popup with multiple buttons
# $buttons = @(
#     @{ Label = "View Status"; Url = "https://status.example.com" },
#     @{ Label = "Contact IT"; Url = "https://helpdesk.example.com" }
# )
# Show-CustomIncidentPopup -Title "System Outage" -Message "The CRM system is currently unavailable. Our team is working to restore service as quickly as possible." -CustomButtons $buttons -Critical