#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

function Get-HKLMRegistryValue
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$KeyPath,
		[Parameter(Mandatory = $false)]
		[string]$ValueName
	)
	
	try
	{
		# Ensure the path starts with HKLM:
		if (-not $KeyPath.StartsWith("HKLM:\"))
		{
			$KeyPath = "HKLM:\$KeyPath"
		}
		
		# Check if the key exists
		if (-not (Test-Path -Path $KeyPath))
		{
			Write-Error "Registry key '$KeyPath' does not exist."
			return $null
		}
		
		# If a specific value is requested
		if ($ValueName)
		{
			$value = Get-ItemProperty -Path $KeyPath -Name $ValueName -ErrorAction SilentlyContinue
			if ($value -ne $null)
			{
				return $value.$ValueName
			}
			else
			{
				Write-Error "Value '$ValueName' does not exist in key '$KeyPath'"
				return $null
			}
		}
		# Return all values in the key
		else
		{
			return Get-ItemProperty -Path $KeyPath
		}
	}
	catch
	{
		Write-Error "Error accessing registry: $_"
		return $null
	}
}

# Example usage:
# Get a specific value from a registry key
# Get-HKLMRegistryValue -KeyPath "SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ValueName "ProductName"

# Get all values from a registry key
# Get-HKLMRegistryValue -KeyPath "SOFTWARE\Microsoft\Windows NT\CurrentVersion"





