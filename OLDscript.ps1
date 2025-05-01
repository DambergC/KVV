# Site configuration

$SiteCode = "KV1" # Site code

$ProviderMachineName = "vntsql0299.kvv.se" # SMS Provider machine name

 

# Customizations

$initParams = @{}

#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging

#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

 

# Do not change anything below this line

 

# Import the ConfigurationManager.psd1 module

if((Get-Module ConfigurationManager) -eq $null) {

    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams

}

 

# Connect to the site's drive if it is not already present

if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {

    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams

}

 

# Set the current location to be the site code.

Set-Location "$($SiteCode):\" @initParams

 

Write-Host "Retrieving CM device objects from CM server by WMI" -ForegroundColor White

$CMDevObjects = (Get-WmiObject -ComputerName vntsql0299.kvv.se -Class SMS_R_SYSTEM -Namespace root\sms\site_KV1 | Where-Object {$_.Name -like "*"})

 

Write-Host "Retrieving DHCP scopes from DHCP server" -ForegroundColor White

$dhcpScopes = Get-DhcpServerv4Scope -ComputerName vntdhcp0002.kvv.se

 

Write-Host "Retrieving MITA computer objects from Active Directory" -ForegroundColor White

$ADComputersMITA = Get-ADComputer -SearchBase "OU=Windows10 SMP,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se" -Filter {Name -like "*"} -properties Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,Enabled,DistinguishedName | Select-Object Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,Enabled,DistinguishedName

 

Write-Host "Retrieving T1 computer objects from Active Directory" -ForegroundColor White

$ADComputersT1 = Get-ADComputer -SearchBase "OU=Windows10 T1,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se" -Filter {Name -like "*"} -properties Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,Enabled,DistinguishedName | Select-Object Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,Enabled,DistinguishedName

 

 

function Get-ScopeName($SubnetIP){

    return ($dhcpScopes | Where-Object {$_.ScopeId -eq $SubnetIP}).Name

 

}

 

function Get-ComputerScopeNames($ComputerName){

    $Subnets = ($CMDevObjects | Where-Object {$_.Name -eq $ComputerName}).IPSubnets

    $Subnets = $Subnets | Where-Object {$_ -notin ("10.0.0.0")}

    $SubnetNames += foreach($Subnet in $Subnets){Get-ScopeName $Subnet}

  

    #return  $SubnetNames | Where-Object {$_ -ne $null}

    return  ($SubnetNames | Where-Object {$_ -ne $null}) -join ", "

}

 

 

 

 

 

 

 

#$dhcpLeases = Get-DhcpServerv4Scope -ComputerName vntdhcp0002.kvv.se | Get-DhcpServerv4Lease -ComputerName vntdhcp0002.kvv.se

$outputComputers = @()

 

 

$MITAs = 0

ForEach($ADComputer in $ADComputersMITA){

    $MITAs += 1

 

    Write-Progress -Activity "MITA" -Status "Running" -PercentComplete ($MITAs/$ADComputersMITA.Length*100)

    $SLSQLQuery = "SELECT SYS.Name0 AS Name,CS.Manufacturer0 AS Manufacturer,CS.Model0 AS Model FROM v_R_System AS SYS JOIN v_GS_COMPUTER_SYSTEM CS ON CS.ResourceID = SYS.ResourceID WHERE SYS.Name0 = '$($ADComputer.Name)'"

    #$SLSQLQuery

    $CMDevice = Get-CMDevice -Name $ADComputer.Name

    $DeviceFQDN = $ADComputer.Name +".kvv.se"

    $dhcpScope = $dhcpLeases | Where-Object HostName -eq ($ADComputer.Name +".kvv.se") | Get-DhcpServerv4Scope -ComputerName vntdhcp0002.kvv.se

    $primaryUsersList = $null

    $currentUsersList = $null

    $adCurrentUser = $null

    $subnetScopesList = $null

    $currentUserLocation = $null

    $cmBoundaryGroups = $null

 

    $SubnetScopes = Get-ComputerScopeNames -ComputerName $ADComputer.Name

 

 

    $primaryUsers = [System.Collections.Generic.List[string]]::new()

    if($CMDevice.PrimaryUser -ne $null){

        ForEach($primaryUser in $CMDevice.PrimaryUser.Split(",")){

            $adPrimaryUser = Get-ADUser $primaryUser.ToUpper().Replace("KVV\","") -Properties Name,Title

            $primaryUsers.Add($adPrimaryUser.Name +" (" +$adPrimaryUser.Title +")")

            $primaryUsersList = $primaryUsers -join ", "

        }

    }

 

 

    $currentUsers = [System.Collections.Generic.List[string]]::new()

    if($CMDevice.UserName -ne $null){

        ForEach($currentUser in $CMDevice.UserName.Split(",")){

            $adCurrentUser = Get-ADUser $currentUser.ToUpper().Replace("KVV\","") -Properties Name,Title

            $currentUsers.Add($adCurrentUser.Name)

            $currentUsersList = $currentUsers -join ", "

        }

    }

 

 

    If($ADComputer.DistinguishedName -like "*Delad enhet*"){

        If($ADComputer.DistinguishedName -like "*Fotostation*"){$Platform = "DMITA Fotostation"}

        Else{$Platform = "DMITA"}

    }

    Else{$Platform = "MITA"}

   

    if($currentUsersList -ne $null){

        $currentUserLocation = $currentUsersList.Substring(($currentUsersList.LastIndexOf("-")+1),($currentUsersList.Length-$currentUsersList.LastIndexOf("-")-1)).Trim()

    }

   

    if($CMDevice.BoundaryGroups -ne $null){

        $cmBoundaryGroups = $CMDevice.BoundaryGroups.Replace("CL - ","")

    }

 

    $ComputerDetails = Invoke-Sqlcmd -Database "CM_KV1" -ServerInstance "vntsql0299.kvv.se" -Query $SLSQLQuery -Verbose -TrustServerCertificate

    $myobj = New-Object -TypeName PSObject

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Datornamn" -Value $ADComputer.Name

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Plattform" -Value $Platform

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Serienummer" -Value $CMDevice.SerialNumber

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Operativsystem" -Value $ADComputer.OperatingSystem

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Operativsystemversion" -Value $ADComputer.OperatingSystemVersion

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Tillverkare" -Value $ComputerDetails.Manufacturer

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Modell" -Value $ComputerDetails.Model

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Inloggningsdatum" -Value $ADComputer.LastLogonDate

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Aktiverad" -Value $ADComputer.Enabled

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Primäranvändare" -Value $primaryUsersList

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Användare" -Value $currentUsersList

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Användartitel" -Value $adCurrentUser.Title

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Användarplats" -Value $currentUserLocation

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "DHCPScope" -Value $SubnetScopes

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Boundarygrupp" -Value $cmBoundaryGroups

    #$myobj

       

    $outputComputers += $myobj

  

    #$myobj | Out-File -FilePath c:\temp\outputcomputers.txt -Encoding utf8 -Append

}

 

 

$T1s = 0

ForEach($ADComputer in $ADComputersT1){

    $T1s += 1

 

    Write-Progress -Activity "T1" -Status "Running" -PercentComplete ($T1s/$ADComputersT1.Length*100)

    $SLSQLQuery = "SELECT SYS.Name0 AS Name,CS.Manufacturer0 AS Manufacturer,CS.Model0 AS Model FROM v_R_System AS SYS JOIN v_GS_COMPUTER_SYSTEM CS ON CS.ResourceID = SYS.ResourceID WHERE SYS.Name0 = '$($ADComputer.Name)'"

    #$SLSQLQuery

    $CMDevice = Get-CMDevice -Name $ADComputer.Name

    $DeviceFQDN = $ADComputer.Name +".kvv.se"

    $dhcpScope = $dhcpLeases | Where-Object HostName -eq ($ADComputer.Name +".kvv.se") | Get-DhcpServerv4Scope -ComputerName vntdhcp0002.kvv.se

    $primaryUsersList = $null

    $currentUsersList = $null

    $adCurrentUser = $null

    $currentUserLocation = $null

    $cmBoundaryGroups = $null

 

    $SubnetScopes = Get-ComputerScopeNames -ComputerName $ADComputer.Name

 

   

    $primaryUsers = [System.Collections.Generic.List[string]]::new()

    if($CMDevice.PrimaryUser -ne $null){

        ForEach($primaryUser in $CMDevice.PrimaryUser.Split(",")){

            $adPrimaryUser = Get-ADUser $primaryUser.ToUpper().Replace("KVV\","") -Properties Name,Title

            $primaryUsers.Add($adPrimaryUser.Name +" (" +$adPrimaryUser.Title +")")

            $primaryUsersList = $primaryUsers -join ", "

        }

    }

 

    $currentUsers = [System.Collections.Generic.List[string]]::new()

    if($CMDevice.CurrentLogonUser -ne $null){

        ForEach($currentUser in $CMDevice.UserName.Split(",")){

            $adCurrentUser = Get-ADUser $currentUser.ToUpper().Replace("KVV\","") -Properties Name,Title

            $currentUsers.Add($adCurrentUser.Name)

            $currentUsersList = $currentUsers -join ", "

        }

    }

 

    If($ADComputer.DistinguishedName -like "*ConfigMgrDP*"){

        $Platform = "Desktop DistP"

    }

    Else{$Platform = "Desktop"}

 

    if($currentUsersList -ne $null){

        $currentUserLocation = $currentUsersList.Substring(($currentUsersList.LastIndexOf("-")+1),($currentUsersList.Length-$currentUsersList.LastIndexOf("-")-1)).Trim()

    }

 

    if($CMDevice.BoundaryGroups -ne $null){

        $cmBoundaryGroups = $CMDevice.BoundaryGroups.Replace("CL - ","")

    }

 

   

    $ComputerDetails = Invoke-Sqlcmd -Database "CM_KV1" -ServerInstance "vntsql0299.kvv.se" -Query $SLSQLQuery -Verbose -TrustServerCertificate

    $myobj = New-Object -TypeName PSObject

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Datornamn" -Value $ADComputer.Name

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Plattform" -Value $Platform

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Serienummer" -Value $CMDevice.SerialNumber

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Operativsystem" -Value $ADComputer.OperatingSystem

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Operativsystemversion" -Value $ADComputer.OperatingSystemVersion

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Tillverkare" -Value $ComputerDetails.Manufacturer

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Modell" -Value $ComputerDetails.Model

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Inloggningsdatum" -Value $ADComputer.LastLogonDate

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Aktiverad" -Value $ADComputer.Enabled

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Primäranvändare" -Value $primaryUsersList

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Användare" -Value $currentUsersList

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Användartitel" -Value $adCurrentUser.Title

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Användarplats" -Value $currentUserLocation

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "DHCPScope" -Value $SubnetScopes

    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Boundarygrupp" -Value $cmBoundaryGroups

    #$myobj

       

    $outputComputers += $myobj

}

 

$outputComputers | Export-Csv -Path C:\temp\computerslocation.csv -Delimiter ";" -NoTypeInformation -Encoding UTF8
