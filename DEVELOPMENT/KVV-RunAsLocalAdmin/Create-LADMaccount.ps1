<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2025 v5.9.252
	 Created on:   	6/21/2025 9:46 AM
	 Created by:   	Administrator
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

<#
.SYNOPSIS
    Installs or uninstalls a local user account named 'LADM' with specified properties.

.DESCRIPTION
    This script creates or removes a local user account named 'LADM' with the password '1DifficultP@ssW0rd'.
    The account is added to the local Administrators group, is set to never expire, and the user cannot change the password.
    Supports install and uninstall actions via -Install or -Uninstall parameters.

.PARAMETER Install
    Creates the 'LADM' account and configures it.

.PARAMETER Uninstall
    Removes the 'LADM' account and its group membership.

.EXAMPLE
    .\Manage-LADMUser.ps1 -Install

.EXAMPLE
    .\Manage-LADMUser.ps1 -Uninstall
#>

<#
.SYNOPSIS
    Installs or uninstalls a local user account named 'LADM' on the local computer only.

.DESCRIPTION
    This script creates or removes a local user account named 'LADM' with the password '1DifficultP@ssW0rd'
    on the local computer only. The account is added to the local Administrators group, is set to never 
    expire, and the user cannot change the password. If the account already exists, the password will be reset.

.EXAMPLE
    # When run as PS1:
    .\Manage-LADMUser.ps1 install

.EXAMPLE
    # When run as EXE:
    Manage-LADMUser.exe install
#>

# Constants
$userName = "LADM"
$password = "1DifficultP@ssW0rd"
$groupName = "Administrators"

function Write-Log
{
	param (
		[string]$Message,
		[string]$Type = "INFO"
	)
	$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	Write-Host "[$timestamp] [$Type] $Message"
}

function Install-LADMUser
{
	Write-Log "Starting installation/configuration of local user '$userName' on computer: $env:COMPUTERNAME"
	$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
	
	# Check if account exists
	try
	{
		$userExists = $null -ne (Get-LocalUser -Name $userName -ErrorAction SilentlyContinue)
		if ($userExists)
		{
			Write-Log "User '$userName' already exists on computer $env:COMPUTERNAME. Resetting password and properties."
			
			# Reset the password
			Set-LocalUser -Name $userName -Password $securePassword -ErrorAction Stop
			Write-Log "Password has been reset for user '$userName'."
			
			# Ensure account properties are set correctly
			Set-LocalUser -Name $userName -PasswordNeverExpires $true -UserMayChangePassword $false -ErrorAction Stop
			Write-Log "Account properties updated for user '$userName'."
		}
		else
		{
			# Create new user
			$userParams = @{
				Name	    = $userName
				Password    = $securePassword
				FullName    = "Local Admin"
				Description = "Local admin account for management"
			}
			
			# Create the user first
			$newUser = New-LocalUser @userParams -ErrorAction Stop
			
			# Then set account properties
			Set-LocalUser -Name $userName -PasswordNeverExpires $true -UserMayChangePassword $false -ErrorAction Stop
			
			Write-Log "User '$userName' created on local computer $env:COMPUTERNAME."
		}
	}
	catch
	{
		Write-Log "Failed to create or update user: $_" "ERROR"
		return
	}
	
	# Add to local Administrators group
	try
	{
		$groupMember = Get-LocalGroupMember -Group $groupName -Member "$env:COMPUTERNAME\$userName" -ErrorAction SilentlyContinue
		if ($null -eq $groupMember)
		{
			Add-LocalGroupMember -Group $groupName -Member "$env:COMPUTERNAME\$userName" -ErrorAction Stop
			Write-Log "User '$userName' added to local '$groupName' group on $env:COMPUTERNAME."
		}
		else
		{
			Write-Log "User '$userName' is already a member of local '$groupName' on $env:COMPUTERNAME."
		}
	}
	catch
	{
		Write-Log "Failed to add user to Administrators group: $_" "ERROR"
	}
	
	Write-Log "Installation/configuration completed on local computer $env:COMPUTERNAME."
}

function Uninstall-LADMUser
{
	Write-Log "Starting uninstallation of local user '$userName' from computer: $env:COMPUTERNAME"
	
	# Remove from Administrators group
	try
	{
		$groupMember = Get-LocalGroupMember -Group $groupName -Member "$env:COMPUTERNAME\$userName" -ErrorAction SilentlyContinue
		if ($null -ne $groupMember)
		{
			Remove-LocalGroupMember -Group $groupName -Member "$env:COMPUTERNAME\$userName" -ErrorAction Stop
			Write-Log "User '$userName' removed from local '$groupName' group on $env:COMPUTERNAME."
		}
	}
	catch
	{
		Write-Log "Note: User either not in group or error removing from group: $_" "WARN"
	}
	
	# Remove user
	try
	{
		$user = Get-LocalUser -Name $userName -ErrorAction SilentlyContinue
		if ($null -ne $user)
		{
			Remove-LocalUser -Name $userName -ErrorAction Stop
			Write-Log "User '$userName' deleted from local computer $env:COMPUTERNAME."
		}
		else
		{
			Write-Log "User '$userName' does not exist on local computer $env:COMPUTERNAME."
		}
	}
	catch
	{
		Write-Log "Failed to remove user: $_" "ERROR"
	}
	
	Write-Log "Uninstallation completed on local computer $env:COMPUTERNAME."
}

# Handling command-line arguments for both PS1 and EXE modes
$action = ""

# Check if running as EXE or PS1
if ($MyInvocation.MyCommand.Name -match '\.exe$')
{
	# EXE mode - use $args array
	if ($args.Count -ge 1)
	{
		$action = $args[0].ToString().ToLower()
	}
}
else
{
	# PS1 mode - check for both named and positional parameters
	if ($args.Count -ge 1)
	{
		$action = $args[0].ToString().ToLower()
	}
}

# Verify we're running with admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin)
{
	Write-Log "This script requires administrative privileges. Please run as administrator." "ERROR"
	exit 1
}

# Process based on action
switch ($action)
{
	"install" {
		Install-LADMUser
	}
	"uninstall" {
		Uninstall-LADMUser
	}
	default {
		Write-Host "Usage:install|uninstall"
		Write-Host "  install   - Creates/configures the LADM account and adds it to local Administrators"
		Write-Host "             - Resets password if the account already exists"
		Write-Host "  uninstall - Removes the LADM account and its local group membership"
	}
}