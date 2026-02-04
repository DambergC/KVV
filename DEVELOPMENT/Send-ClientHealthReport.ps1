
<#
	.SYNOPSIS
		HTML document status Clienthealth
	
	.DESCRIPTION
		The script creates a HMTL document with status on Clienthealth based on what is registred in Clienthealth database.
	
	.EXAMPLE
		PS C:\> .\Send-ClientHealthReport.ps1
	
	.NOTES
		Additional information about the file.
#>


[System.Xml.XmlDocument]$xml = Get-Content .\ScriptConfigTEST.xml

$Logfilepath = $xml.Configuration.Logfile.Path
$logfilename = $xml.Configuration.Logfile.ClientHealthLogName
$Logfile = $Logfilepath + $logfilename
$Logfilethreshold = $xml.Configuration.Logfile.Logfilethreshold

$HTMLFileSavePath = $xml.Configuration.HTMLfilePath
$HTMLFilename = $xml.Configuration.ClientHealthHTMLName

$scriptname = $MyInvocation.MyCommand.Name
$siteserver = $xml.Configuration.SiteServer
$filedate = get-date -Format yyyMMdd
$SMTP = $xml.Configuration.MailSMTP
$MailFrom = $xml.Configuration.Mailfrom
$MailPortnumber = $xml.Configuration.MailPort
$MailCustomer = $xml.Configuration.MailCustomer


function Rotate-Log
{
	$target = Get-ChildItem $Logfilepath -Filter "Client*.log"
	$datetime = Get-Date -uformat "%Y-%m-%d-%H%M"
	
	$target | ForEach-Object {
		
		if ($_.Length -ge $Logfilethreshold)
		{
			Write-Host "file named $($_.name) is bigger than $Logfilethreshold B"
			$newname = "$($_.BaseName)_${datetime}.log"
			Rename-Item $_.fullname $newname
			
			if (test-path "$Logfilepath\OLDLOG")
			{
				Move-Item .\logfiles\$newname -Destination "$Logfilepath\OLDLOG"
				Write-Host "Done rotating file"
			}
			
			else
			{
				new-item -Path $Logfilepath -Name OLDLOG -ItemType Directory
				Move-Item .\logfiles\$newname -Destination "$Logfilepath\OLDLOG"
				Write-Host "Done rotating file"
			}
		}
		
		else
		{
			Write-Host "file named $($_.name) is not bigger than $Logfilethreshold B"
		}
		Write-Host "Logfile checked!"
	}
}

Rotate-Log

function Write-Log
{
	Param ([string]$LogString)
	$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
	$LogMessage = "$Stamp $LogString"
	Add-content $LogFile -value $LogMessage
}

# Send-MailkitMessage - https://github.com/austineric/Send-MailKitMessage
if (-not (Get-Module -name send-mailkitmessage))
{
	#Install-Module send-mailkitmessage -ErrorAction SilentlyContinue
	Import-Module send-mailkitmessage
}

# pswritehtml - https://github.com/EvotecIT/PSWriteHTML
if (-not (Get-Module -name PSWriteHTML))
{
	#Install-Module PSWriteHTML -ErrorAction SilentlyContinue
	Import-Module PSWriteHTML
}


#########################################################
# Section to extract monthname, year and weeknumbers
#########################################################

$ResultDiskspaceServers = @()
$ResultLastbootTimeServers = @()
$ResultDiskspaceClients = @()
$ResultLastbootTimeClients = @()
$ResultStatWindowsBuild = @()

$todayDefault = Get-Date
$todayshort = $todayDefault.ToShortDateString()


$TitleDate = get-date -DisplayHint Date
$htmlFile = $HTMLFileSavePath + $HTMLFilename + "_" + $todayshort + ".html"

## Server start ##
$queryDiskspaceServers = "SELECT Hostname,OperatingSystem,OSDiskFreeSpace FROM dbo.Clients WHERE OperatingSystem LIKE '%server%' AND OSDiskFreeSpace < '20' ORDER BY OSDiskFreeSpace ASC"
$dataDiskspaceServers = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryDiskspaceServers

foreach ($res in $dataDiskspaceServers)
{
	$object = New-Object -TypeName PSObject
	$object | Add-Member -MemberType NoteProperty -Name 'Servername' -Value $res.Hostname
	$object | Add-Member -MemberType NoteProperty -Name 'OS' -Value $res.OperatingSystem
	$object | Add-Member -MemberType NoteProperty -Name 'Free Disk Space' -Value $res.osdiskfreespace
	$ResultDiskspaceServers += $object
}

$queryLastbootTimeServers = "SELECT Hostname,OperatingSystem,LastBootTime FROM dbo.Clients WHERE OperatingSystem like '%server%' AND LastbootTime <= DATEADD(day, -20, CAST(getdate() As date)) ORDER BY LastBootTime ASC"
$DataLastbootTimeServers = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryLastbootTimeServers

foreach ($res in $DataLastbootTimeServers)
{
	$object = New-Object -TypeName PSObject
	$object | Add-Member -MemberType NoteProperty -Name 'Servername' -Value $res.Hostname
	$object | Add-Member -MemberType NoteProperty -Name 'OS' -Value $res.OperatingSystem
	$object | Add-Member -MemberType NoteProperty -Name 'LastBootTime' -Value $res.LastbootTime
	$ResultLastbootTimeServers += $object
}

## Server end ##

## Client start ##
$queryDiskspaceClients = "SELECT TOP 20 Hostname,LastloggedonUser,OperatingSystem,OSDiskFreeSpace,hwinventory FROM dbo.Clients WHERE OperatingSystem LIKE '%Enterprise%' AND Hwinventory >= DATEADD(day, -1, CAST(getdate() As date)) AND OSDiskFreeSpace < '20' ORDER BY OSDiskFreeSpace ASC"
$dataDiskspaceClients = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryDiskspaceClients

foreach ($res in $dataDiskspaceClients)
{
	$object = New-Object -TypeName PSObject
	$object | Add-Member -MemberType NoteProperty -Name 'Servername' -Value $res.Hostname
	$object | Add-Member -MemberType NoteProperty -Name 'LastUser' -Value $res.LastLoggedonUser
	$object | Add-Member -MemberType NoteProperty -Name 'OS' -Value $res.OperatingSystem
	$object | Add-Member -MemberType NoteProperty -Name 'Free Disk Space' -Value $res.osdiskfreespace
	$object | Add-Member -MemberType NoteProperty -Name 'HwInventory' -Value $res.hwinventory
	$ResultDiskspaceClients += $object
}

$queryLastbootTimeClients = "SELECT TOP 20 Hostname,OperatingSystem,LastBootTime,HWInventory,LastloggedonUser FROM dbo.Clients WHERE OperatingSystem LIKE '%Enterprise%' AND Hwinventory >= DATEADD(day, -1, CAST(getdate() As date)) ORDER BY LastBootTime ASC"
$DataLastbootTimeClients = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryLastbootTimeClients

foreach ($res in $DataLastbootTimeClients)
{
	$object = New-Object -TypeName PSObject
	$object | Add-Member -MemberType NoteProperty -Name 'Servername' -Value $res.Hostname
	$object | Add-Member -MemberType NoteProperty -Name 'LastUser' -Value $res.LastLoggedonUser
	$object | Add-Member -MemberType NoteProperty -Name 'OS' -Value $res.OperatingSystem
	$object | Add-Member -MemberType NoteProperty -Name 'LastBootTime' -Value $res.LastbootTime
	$object | Add-Member -MemberType NoteProperty -Name 'HwInventory' -Value $res.hwinventory
	$ResultLastbootTimeClients += $object
}

## Client end ##

## Statistics start ##

$queryStatOSVersionServers = "SELECT Operatingsystem, COUNT(Operatingsystem) FROM dbo.Clients WHERE OperatingSystem LIKE '%server%' GROUP BY OperatingSystem ORDER BY OperatingSystem"
$DataStatOSVersionServers = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryStatOSVersionServers

$queryStatOSVersionClients = "SELECT Operatingsystem, COUNT(Operatingsystem) FROM dbo.Clients WHERE OperatingSystem LIKE '%Enterprise%' GROUP BY OperatingSystem ORDER BY OperatingSystem"
$DataStatOSVersionClients = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryStatOSVersionClients


$queryStatWindowsBuild = "SELECT Build, COUNT(Build) FROM dbo.Clients WHERE OperatingSystem LIKE '%10 Enterprise%' GROUP BY build ORDER BY build"
$DataStatWindowsBuild = Invoke-Sqlcmd -ServerInstance $siteserver -Database ClientHealth -Query $queryStatWindowsBuild

foreach ($res in $DataStatWindowsBuild)
{
	$object = New-Object -TypeName PSObject
	$object | Add-Member -MemberType NoteProperty -Name 'Build' -Value $res.Build
	$object | Add-Member -MemberType NoteProperty -Name 'Count' -Value $res.Column1
	
	$ResultStatWindowsBuild += $object
}



#region Script part 2 Create the html-file to be distributed

New-HTML -TitleText "ClientHealth Report - Kriminalvården" -FilePath $htmlFile -ShowHTML -Online {
	
	New-HTMLHeader {
		New-HTMLSection -Invisible {
			New-HTMLPanel -Invisible {
				New-HTMLText -Text "Kriminalvården - ClientHealth Report" -FontSize 35 -Color Darkblue -FontFamily Arial -Alignment center
				New-HTMLHorizontalLine
			}
		}
	}
	
	New-HTMLSection -HeaderTextSize 20 -HeaderBackGroundColor black -Title "ClientHealth - Servers"{
		
		New-HTMLTable -Title 'Diskspace' -DisableSearch -DataTable $ResultDiskspaceServers -PagingLength 35 -Style compact
		New-HTMLTable -Title 'Boottime' -DisableSearch -DataTable $ResultLastbootTimeServers -PagingLength 35 -Style compact
		
	}
	
	New-HTMLSection -HeaderTextSize 20 -HeaderBackGroundColor black -Title "ClientHealth - Clients"{
		
		New-HTMLTable -Title 'Diskspace' -DisableSearch -DataTable $ResultDiskspaceClients -PagingLength 35 -Style compact
		New-HTMLTable -Title 'Boottime' -DisableSearch -DataTable $ResultLastbootTimeClients -PagingLength 35 -Style compact
		
	}
	
	New-HTMLSection -HeaderTextSize 20 -HeaderBackGroundColor black -Title "ClientHealth - Statistics"{
		
		New-HTMLChart -Gradient -Title 'Servers' {
			
			foreach ($result in $DataStatOSVersionServers)
			{
				New-ChartPie -Name $result.operatingsystem -Value $result.column1
				
				
			}
		}
		
		New-HTMLChart -Gradient -Title 'Clients' {
			
			foreach ($result in $DataStatOSVersionClients)
			{
				New-ChartPie -Name $result.operatingsystem -Value $result.column1
				
				
			}
		}
		
		
	}
	
	New-HTMLSection -HeaderTextSize 20 -HeaderBackGroundColor black -Title "ClientHealth - Statistics part 2"{
		
		New-HTMLChart -Gradient -Title 'Windows Versions' {
			
			foreach ($result in $ResultStatWindowsBuild)
			{
				New-ChartPie -Name $result.Build -Value $result.count
				
				
			}
			
			
		}
	}
	
	
	
	New-HTMLFooter {
		
		New-HTMLSection -Invisible {
			
			New-HTMLPanel -Invisible {
				New-HTMLHorizontalLine
				New-HTMLText -Text "Denna lista skapades $todaydefault" -FontSize 20 -Color Darkblue -FontFamily Arial -Alignment center -FontStyle italic
			}
			
		}
	}
	
	
}


#endregion
