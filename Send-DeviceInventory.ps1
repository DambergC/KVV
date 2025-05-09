
<#
	===========================================================================
	Values needed to be updated before running the script
	===========================================================================
#>
$scriptversion = '1.0'
$scriptname = $MyInvocation.MyCommand.Name

[System.Xml.XmlDocument]$xml = Get-Content .\ScriptConfigDevInvent.xml

$siteserver = $xml.Configuration.SiteServer
$filedate = get-date -Format yyyMMdd
$HTMLFileSavePath = "G:\Scripts\Outfiles\DeviceInvent_$filedate.HTML"
$CSVFileSavePath = "G:\Scripts\Outfiles\DeviceInvent_$filedate.csv"
$SMTP = $xml.Configuration.MailSMTP
$MailFrom = $xml.Configuration.Mailfrom
$MailPortnumber = $xml.Configuration.MailPort
$MailCustomer = $xml.Configuration.MailCustomer
$logfilename = $xml.Configuration.Logfile.Name
$logfilePath = $xml.Configuration.Logfile.Path
$logfile = $logfilePath+"\"+$logfilename

$today = get-date -Format yyyy-MM-dd

function Write-Log
{
Param ([string]$LogString)
$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $LogFile -value $LogMessage
}

function Rotate-log 
{
    $target = Get-ChildItem $Logfile -Filter "windows*.log"
    $threshold = 30
    $datetime = Get-Date -uformat "%Y-%m-%d-%H%M"
    $target | ForEach-Object {
    if ($_.Length -ge $threshold) { 
        Write-Host "file named $($_.name) is bigger than $threshold KB"
        $newname = "$($_.BaseName)_${datetime}.log_old"
        Rename-Item $_.fullname $newname
        Write-Host "Done rotating file" 
    }
    else{
         Write-Host "file named $($_.name) is not bigger than $threshold KB"
    }
    Write-Host " "
}
}



<#
	===========================================================================
	Powershell modules needed in the script
	===========================================================================

	Send-MailkitMessage - https://github.com/austineric/Send-MailKitMessage

	pswritehtml - https://github.com/EvotecIT/PSWriteHTML

	PatchManagementSupportTools - Created by Christian Damberg, Cygate
	https://github.com/DambergC/PatchManagement/tree/main/PatchManagementSupportTools

	DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!
#>

#region modules

if (-not (Get-Module -name send-mailkitmessage))
{
	#Install-Module send-mailkitmessage -ErrorAction SilentlyContinue
	Import-Module send-mailkitmessage
}


if (-not (Get-Module -name PSWriteHTML))
{
	#Install-Module PSWriteHTML -ErrorAction SilentlyContinue
	Import-Module PSWriteHTML
}

if (-not (Get-Module -name PatchManagementSupportTools))
{
	#Install-Module PatchManagementSupportTools -ErrorAction SilentlyContinue
	Import-Module PatchManagementSupportTools
}

#endregion


<#
	===========================================================================		
	Date-section

	DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!
	===========================================================================
#>







$TitleDate = get-date -DisplayHint Date
$counter = 0

#check if script should run or not


#Region Script part 1 collect info from selected collection and check devices membership in Collections with Maintenance Windows


$query = "SELECT
    SYS.Netbios_Name0 AS 'Device_Name',
    BIOS.SerialNumber0 AS 'Serial_Number',
    CS.Model0 AS 'Model',
    CS.Manufacturer0 AS 'Manufacturer',
    OS.Caption0 AS 'Operating_System',
    OS.Version0 AS 'OS_Version',
    OS.BuildNumber0 AS 'Build_Number',
    CASE 
        WHEN CS.Manufacturer0 LIKE '%VMware%' THEN 'Virtual'
        WHEN CS.Manufacturer0 LIKE '%Microsoft%' AND CS.Model0 LIKE '%Virtual%' THEN 'Virtual'
        WHEN CS.Manufacturer0 LIKE '%VirtualBox%' THEN 'Virtual'
        ELSE 'Physical'
    END AS 'Device Type',
    STRING_AGG(vru.Name0, ', ') AS 'Primary_Users',
    SYS.User_Name0 AS 'Last_Logon_User',
    SYS.Last_Logon_Timestamp0 AS 'Last_Logon_Time', -- Added Last_Logon_Time
    SYS.Resource_Domain_OR_Workgr0 AS 'Domain',
    STRING_AGG(vru.Mail0, '; ') AS 'User_Emails',
    IPADDR.IP_Addresses0 AS 'IPv4_Address',
    bginfo.BoundaryName AS 'Boundary_Name',
    bginfo.BoundaryValue AS 'Boundary_Value',
    bginfo.BoundaryGroupName AS 'Boundary_Group_Name',
    SCCM.OUpath00 AS 'OU_Path' -- Included OUpath00
FROM
    v_R_System AS SYS
JOIN
    v_GS_COMPUTER_SYSTEM AS CS
    ON SYS.ResourceID = CS.ResourceID
JOIN
    v_GS_PC_BIOS AS BIOS
    ON SYS.ResourceID = BIOS.ResourceID
JOIN
    v_GS_OPERATING_SYSTEM AS OS
    ON SYS.ResourceID = OS.ResourceID
LEFT JOIN 
    v_UsersPrimaryMachines upm 
    ON SYS.ResourceID = upm.MachineID
LEFT JOIN 
    v_R_User vru 
    ON upm.UserResourceID = vru.ResourceID
LEFT JOIN
    v_R_User LLT -- Joining with LastLogon table
    ON SYS.ResourceID = LLT.ResourceID
LEFT JOIN 
    dbo.SCCMItems64_DATA AS SCCM -- Joining with SCCMItems64_DATA table
    ON SYS.ResourceID = SCCM.MachineID -- Correctly connecting MachineID to ResourceID
OUTER APPLY (
    SELECT TOP 1 IP.IP_Addresses0 
    FROM v_RA_System_IPAddresses IP 
    WHERE IP.ResourceID = SYS.ResourceID 
    AND IP.IP_Addresses0 NOT LIKE '%:%' -- Filter out IPv6
    AND IP.IP_Addresses0 NOT LIKE '169.254.%' -- Filter out APIPA addresses
    ORDER BY 
        CASE WHEN IP.IP_Addresses0 LIKE '10.%' THEN 1
             WHEN IP.IP_Addresses0 LIKE '172.1[6-9].%' OR IP.IP_Addresses0 LIKE '172.2[0-9].%' OR IP.IP_Addresses0 LIKE '172.3[0-1].%' THEN 2
             WHEN IP.IP_Addresses0 LIKE '192.168.%' THEN 3
             ELSE 4 END, -- Prioritize private IPs
        IP.IP_Addresses0
) AS IPADDR
OUTER APPLY (
    SELECT DISTINCT 
        bg.GroupID,
        bg.Name AS 'BoundaryGroupName',
        b.DisplayName AS 'BoundaryName',
        b.Value AS 'BoundaryValue'
    FROM 
        vSMS_Boundary b
    JOIN 
        vSMS_BoundaryGroupMembers bgm ON b.BoundaryID = bgm.BoundaryID
    JOIN 
        vSMS_BoundaryGroup bg ON bgm.GroupID = bg.GroupID
    WHERE
        b.BoundaryType = 3 AND -- IP Address Range
        (CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 4)) * 16777216) +
        (CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 3)) * 65536) +
        (CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 2)) * 256) +
         CONVERT(BIGINT, PARSENAME(IPADDR.IP_Addresses0, 1))
        BETWEEN
        (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 4)) * 16777216) +
        (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 3)) * 65536) +
        (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 2)) * 256) +
         CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, 1, CHARINDEX('-', b.Value) - 1), 1))
        AND
        (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 4)) * 16777216) +
        (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 3)) * 65536) +
        (CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 2)) * 256) +
         CONVERT(BIGINT, PARSENAME(SUBSTRING(b.Value, CHARINDEX('-', b.Value) + 1, LEN(b.Value)), 1))
) AS bginfo
WHERE
    CASE 
        WHEN CS.Manufacturer0 LIKE '%VMware%' THEN 'Virtual'
        WHEN CS.Manufacturer0 LIKE '%Microsoft%' AND CS.Model0 LIKE '%Virtual%' THEN 'Virtual'
        WHEN CS.Manufacturer0 LIKE '%VirtualBox%' THEN 'Virtual'
        ELSE 'Physical'
    END = 'Physical' -- Filter for physical machines
    AND (OS.Caption0 LIKE '%Windows 10%' OR OS.Caption0 LIKE '%Windows 11%') -- Include only Windows 10 and Windows 11
    AND OS.Caption0 NOT LIKE '%Server%' -- Exclude servers
    AND SYS.Netbios_Name0 NOT LIKE 'WSDP%' -- Exclude devices starting with WSDP
    AND bginfo.BoundaryGroupName NOT LIKE 'CL - Server Net%' -- Exclude specific boundary names
GROUP BY
    SYS.Netbios_Name0,
    BIOS.SerialNumber0,
    CS.Model0,
    CS.Manufacturer0,
    OS.Caption0,
    OS.Version0,
    OS.BuildNumber0,
    SYS.User_Name0,
    sys.Last_Logon_Timestamp0, -- Grouping by Last_Logon_Time
    SYS.Resource_Domain_OR_Workgr0,
    SYS.ResourceID,
    IPADDR.IP_Addresses0,
    bginfo.BoundaryName,
    bginfo.BoundaryValue,
    bginfo.BoundaryGroupName,
    SCCM.OUpath00, -- Grouping by OUpath00
    CASE 
        WHEN CS.Manufacturer0 LIKE '%VMware%' THEN 'Virtual'
        WHEN CS.Manufacturer0 LIKE '%Microsoft%' AND CS.Model0 LIKE '%Virtual%' THEN 'Virtual'
        WHEN CS.Manufacturer0 LIKE '%VirtualBox%' THEN 'Virtual'
        ELSE 'Physical'
    END
ORDER BY
    SYS.Netbios_Name0;"
                          
                          
                          
            
            		$data = Invoke-Sqlcmd -ServerInstance $siteserver -Database CM_KV1 -Query $query -Verbose


$resultColl = @()

foreach ($row in $data)
{
    $object = New-Object -TypeName PSObject
    $object | Add-Member -MemberType NoteProperty -Name 'Device Name' -Value $row.device_name
    $object | Add-Member -MemberType NoteProperty -Name 'Serial Number' -Value $row.Serial_number
    $object | Add-Member -MemberType NoteProperty -Name 'Model' -Value $row.Model
    #$object | Add-Member -MemberType NoteProperty -Name 'Manufacturer' -Value $row.Manufacturer
    $object | Add-Member -MemberType NoteProperty -Name 'Operating System' -Value $row.operating_system
    #$object | Add-Member -MemberType NoteProperty -Name 'OS Version' -Value $row.OS_Version
    $object | Add-Member -MemberType NoteProperty -Name 'Build Number' -Value $row.Build_Number
    #$object | Add-Member -MemberType NoteProperty -Name 'Device Type' -Value $row.'Device Type'
    $object | Add-Member -MemberType NoteProperty -Name 'Last Logon User' -Value $row.Last_logon_user
    $object | Add-Member -MemberType NoteProperty -Name 'Last Logon Time' -Value $row.Last_logon_Time
    #$object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $row.Domain
    #$object | Add-Member -MemberType NoteProperty -Name 'User Email(s)' -Value $row.User_Emails
    $object | Add-Member -MemberType NoteProperty -Name 'IPv4 Address' -Value $row.IPv4_Address
    $object | Add-Member -MemberType NoteProperty -Name 'Boundary Name' -Value $row.Boundary_Name
    $object | Add-Member -MemberType NoteProperty -Name 'Boundary Group Name' -Value $row.Boundary_Group_Name
    #$object | Add-Member -MemberType NoteProperty -Name 'Boundary Value' -Value $row.Boundary_Value
    $object | Add-Member -MemberType NoteProperty -Name 'Primary User(s)' -Value $row.Primary_users

    $object | Add-Member -MemberType NoteProperty -Name 'Type' -Value $row.OU_path

    $resultColl += $object
}


#endregion

#region Script part 2 Create the html-file to be distributed

# New-HTML -TitleText "Patchfönster- Kriminalvården" -FilePath $HTMLFileSavePath -ShowHTML -Online

New-HTML -TitleText "Deviceinventering - Kriminalvården" -Online -FilePath $HTMLFileSavePath -ShowHTML {
	
	New-HTMLHeader {
		New-HTMLSection -HeaderText 'Kriminalvården IT Arbetsplats' -HeaderTextSize 35 -BackgroundColor lightblue -HeaderTextColor Darkblue {
			New-HTMLPanel -BackgroundColor lightblue {
				New-HTMLText -Text "Status Fysiska datorer $today från databasen för ConfigMgr" -FontSize 25 -Color Darkblue -FontFamily Arial -Alignment center -BackGroundColor lightblue
				New-HTMLHorizontalLine
			}
		}
	}
	
	New-HTMLSection -Title "Sorterings- och exportbar data" -HeaderTextSize 20 -HeaderTextColor darkblue{
		
		New-HTMLTable -DataTable $resultColl -PagingLength 35 -AutoSize -Style nowrap
		
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



#Region CSS and HTML for mail thru Send-MailKitMessage

$content = Get-Content -Path $HTMLFileSavePath

#endregion

#Region HTML Mail

$Body = @"

<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>SFysiska datorer ConfigMgr - Kriminalvården</title>
<style>

    th {

        font-family: Arial, Helvetica, sans-serif;
        color: White;
        font-size: 12px;
        border: 1px solid black;
        padding: 3px;
        background-color: Black;

    } 
    p {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 12px;

    } 
    ol {

        font-family: Arial, Helvetica, sans-serif;
        list-style-type: square;
        color: black;
        font-size: 12px;

    }
	    H1 {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 18px;

    }
    tr {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 11px;
        vertical-align: text-top;

    } 

    body {
        background-color: lightgray;
      }
      table {
        border: 1px solid black;
        border-collapse: collapse;
      }

      td {
        border: 1px solid black;
        padding: 5px;
        background-color: #E0F3F7;
      }

</style>
</head>

<body>
	<p><h1>Inventering - fysiska klienter - Kriminalvården</h1></p> 
	<p>Biforgade filer innehåller dagsaktuell data from CM_KV1 databasen kopplad till ConfigMgr.<br><br>
<hr>
</p> 
	<p>Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>

	
	
	
</body>
</html>
 

"@



#endregion

#Region Mailsettings


#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable = $false

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential = [System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer = $SMTP

#port ([int], required)
$Port = $MailPortnumber

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From = [MimeKit.MailboxAddress]$MailFrom

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList = [MimeKit.InternetAddressList]::new()
    
    $recipientlistXML = $xml.Configuration.Recipients | ForEach-Object {$_.Recipients.Email}
    
    foreach ($Recipient in $recipientlistXML)
    
        {
            $RecipientList.Add([MimeKit.InternetAddress]$Recipient)
        }
#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$CCList=[MimeKit.InternetAddressList]::new()
#$CCList.Add([MimeKit.InternetAddress]$EmailToCC)



#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList = [MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")


#subject ([string], required)
$Subject = [string]"Serverpatchning $MailCustomer $monthname $year"

#text body ([string], optional)
#$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody = [string]$Body

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList = [System.Collections.Generic.List[string]]::new()
$AttachmentList.Add("$HTMLFileSavePath")
#$AttachmentList.Add("$CSVFileSavePath")

# Mailparameters
$Parameters = @{
	"UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable
	#"Credential"=$Credential
	"SMTPServer"					 = $SMTPServer
	"Port"						     = $Port
	"From"						     = $From
	"RecipientList"				     = $RecipientList
	#"CCList"=$CCList
	#"BCCList"=$BCCList
	"Subject"					     = $Subject
	#"TextBody"=$TextBody
	"HTMLBody"					     = $HTMLBody
	"AttachmentList"				 = $AttachmentList
}

#endregion

#Region Send Mail

Send-MailKitMessage @Parameters
Write-Log -LogString "$scriptname - Mail on it´s way to $RecipientList"
set-location $PSScriptRoot
Write-Log -LogString "$scriptname - Script exit!"
#endregion
