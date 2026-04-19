Import-Module ActiveDirectory -ErrorAction Stop

$ErrorActionPreference = 'Stop'

# =========================================
# Konfiguration
# Aktiva-datorer-under-InaktivaOU
# =========================================

$searchBase = "OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
$emailRecipients = "HK.IT.CMKalendern@kriminalvarden.se"
$emailFrom       = "noreply@kriminalvarden.se"
$emailSubject    = "Aktiva datorer under Inaktiva"
$smtpServer      = "smtp.kvv.se"
$inactiveRootOU = "OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

# Återflyttsmappning
$returnMoveMap = @(
    [pscustomobject]@{
        InactiveOU = "OU=T1,OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
        TargetOU   = "OU=T1,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
    },
    [pscustomobject]@{
        InactiveOU = "OU=Mita,OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
        TargetOU   = "OU=MITA,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
    },
    [pscustomobject]@{
        InactiveOU = "OU=DeladMita,OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
        TargetOU   = "OU=DMITA,OU=MITA,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"
    }
)
 
# =========================================
# Funktioner
# =========================================
 
function Convert-ToHtmlSafe {
    param(
        [AllowNull()]
        [object]$Value
    )
 
    if ($null -eq $Value) {
        return ""
    }
 
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}
 
function Get-ParentOUFromDN {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistinguishedName
    )
 
    return ($DistinguishedName -replace '^[^,]+,', '')
}
 
function Get-ReturnTargetOU {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistinguishedName,
 
        [Parameter(Mandatory = $true)]
        [array]$MoveMap
    )
 
    $parentOU = Get-ParentOUFromDN -DistinguishedName $DistinguishedName
 
    foreach ($mapping in $MoveMap) {
        if ($parentOU -ieq $mapping.InactiveOU) {
            return $mapping.TargetOU
        }
    }
 
    return $null
}
 
function Send-ResultMail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BodyHtml
    )
 
    Send-MailMessage `
        -From $emailFrom `
        -To $emailRecipients `
        -Subject $emailSubject `
        -Body $BodyHtml `
        -BodyAsHtml `
        -Encoding UTF8 `
        -SmtpServer $smtpServer `
        -ErrorAction Stop
}
 
# =========================================
# Huvudlogik
# =========================================
 
$runTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
 
$movedBackResults = @()
$rootLevelResults = @()
$failedResults    = @()
 
try {
    # Validera OU:er
    Get-ADOrganizationalUnit -Identity $inactiveRootOU -ErrorAction Stop | Out-Null
 
    foreach ($mapping in $returnMoveMap) {
        Get-ADOrganizationalUnit -Identity $mapping.InactiveOU -ErrorAction Stop | Out-Null
        Get-ADOrganizationalUnit -Identity $mapping.TargetOU   -ErrorAction Stop | Out-Null
    }
 
    # Hämta aktiva datorer med LastLogonDate under hela Win11\Inaktiva
    $enabledInactiveComputers = Get-ADComputer `
        -Filter 'Enabled -eq $true' `
        -SearchBase $searchBase `
        -SearchScope Subtree `
        -Properties CanonicalName, LastLogonDate, DistinguishedName, DNSHostName, sAMAccountName `
        -ErrorAction Stop |
        Where-Object { $_.LastLogonDate } |
        Sort-Object Name
 
    foreach ($computer in $enabledInactiveComputers) {
        $parentOU = Get-ParentOUFromDN -DistinguishedName $computer.DistinguishedName
        $targetOU = Get-ReturnTargetOU -DistinguishedName $computer.DistinguishedName -MoveMap $returnMoveMap
 
        # Dator direkt under root-OU -> ingen automatisk återflytt
        if ($parentOU -ieq $inactiveRootOU) {
            $rootLevelResults += [pscustomobject]@{
                Name           = $computer.Name
                DNSHostName    = $computer.DNSHostName
                SamAccountName = $computer.sAMAccountName
                LastLogonDate  = $computer.LastLogonDate
                CanonicalName  = $computer.CanonicalName
                Status         = "Aktiv dator direkt under Win11\Inaktiva"
                ResultMessage  = "Ingen automatisk återflytt. Kräver manuell kontroll."
            }
 
            continue
        }
 
        # Matchar särskild OU för återflytt
        if (-not [string]::IsNullOrWhiteSpace($targetOU)) {
            try {
                Move-ADObject -Identity $computer.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
 
                $movedBackResults += [pscustomobject]@{
                    Name           = $computer.Name
                    DNSHostName    = $computer.DNSHostName
                    SamAccountName = $computer.sAMAccountName
                    LastLogonDate  = $computer.LastLogonDate
                    CanonicalName  = $computer.CanonicalName
                    Status         = "Återflyttad"
                    ResultMessage  = "Datorobjektet var aktivt och flyttades tillbaka till rätt OU: $targetOU"
                }
            }
            catch {
                $failedResults += [pscustomobject]@{
                    Name           = $computer.Name
                    DNSHostName    = $computer.DNSHostName
                    SamAccountName = $computer.sAMAccountName
                    LastLogonDate  = $computer.LastLogonDate
                    CanonicalName  = $computer.CanonicalName
                    Status         = "Fel"
                    ResultMessage  = "Flytt misslyckades: $($_.Exception.Message)"
                }
            }
 
            continue
        }
 
        # Aktiva datorer i andra under-OU:er under Win11\Inaktiva
        $rootLevelResults += [pscustomobject]@{
            Name           = $computer.Name
            DNSHostName    = $computer.DNSHostName
            SamAccountName = $computer.sAMAccountName
            LastLogonDate  = $computer.LastLogonDate
            CanonicalName  = $computer.CanonicalName
            Status         = "Aktiv dator i under-OU utan återflyttsmappning"
            ResultMessage  = "Ingen automatisk återflytt eftersom ingen matchande mål-OU definierades."
        }
    }
 
    $totalFound     = @($enabledInactiveComputers).Count
    $movedBackCount = @($movedBackResults).Count
    $rootLevelCount = @($rootLevelResults).Count
    $failedCount    = @($failedResults).Count
 
    $allRows = @()
    $allRows += $movedBackResults
    $allRows += $rootLevelResults
    $allRows += $failedResults
 
    $tableRows = foreach ($item in $allRows) {
        $lastLogon = if ($item.LastLogonDate) {
            ([datetime]$item.LastLogonDate).ToString("yyyy-MM-dd HH:mm:ss")
        }
        else {
            ""
        }
 
@"
<tr>
    <td>$(Convert-ToHtmlSafe $item.Name)</td>
    <td>$(Convert-ToHtmlSafe $item.DNSHostName)</td>
    <td>$(Convert-ToHtmlSafe $item.SamAccountName)</td>
    <td>$(Convert-ToHtmlSafe $lastLogon)</td>
    <td>$(Convert-ToHtmlSafe $item.CanonicalName)</td>
    <td>$(Convert-ToHtmlSafe $item.Status)</td>
    <td>$(Convert-ToHtmlSafe $item.ResultMessage)</td>
</tr>
"@
    }
 
    if (-not $tableRows) {
        $tableRows = @"
<tr>
    <td colspan="7">Inga aktiva datorobjekt med LastLogonDate hittades under $searchBase.</td>
</tr>
"@
    }
 
    $bodyHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>$emailSubject</title>
    <style>
        body {
            font-family: Segoe UI, Arial, Helvetica, sans-serif;
            font-size: 10pt;
            color: #222222;
        }
        h1 {
            font-size: 16pt;
            color: #1f4e79;
        }
        p {
            margin-bottom: 12px;
        }
        .summary {
            padding: 10px;
            background-color: #f3f6fa;
            border: 1px solid #d9e2f0;
            margin-bottom: 15px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th {
            background-color: #1f4e79;
            color: #ffffff;
            text-align: left;
            padding: 8px;
            border: 1px solid #d9d9d9;
        }
        td {
            padding: 8px;
            border: 1px solid #d9d9d9;
            vertical-align: top;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .footer {
            margin-top: 20px;
            font-size: 9pt;
            color: #666666;
        }
    </style>
</head>
<body>
    <h1>Aktiva datorobjekt under OU=Win11,OU=Inaktiva</h1>
 
    <p>
        Skriptet har kontrollerat aktiva datorobjekt med <b>LastLogonDate</b> under
        <b>OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se</b>.
    </p>
 
    <div class="summary">
        <b>Antal aktiva datorobjekt med LastLogonDate:</b> $totalFound<br>
        <b>Återflyttade datorobjekt:</b> $movedBackCount<br>
        <b>Aktiva datorobjekt som kräver kontroll:</b> $rootLevelCount<br>
        <b>Fel:</b> $failedCount<br>
        <b>Sökbas:</b> $(Convert-ToHtmlSafe $searchBase)<br>
        <b>Tidpunkt:</b> $(Convert-ToHtmlSafe $runTime)
    </div>
 
    <p>
        Datorobjekt som ligger direkt under <b>OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se</b>
        flyttas inte automatiskt tillbaka utan redovisas i rapporten för manuell kontroll.
    </p>
 
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>DNSHostName</th>
                <th>SamAccountName</th>
                <th>LastLogonDate</th>
                <th>CanonicalName</th>
                <th>Status</th>
                <th>ResultMessage</th>
            </tr>
        </thead>
        <tbody>
            $($tableRows -join "`r`n")
        </tbody>
    </table>
 
    <div class="footer">
        Detta mejl genererades automatiskt.
    </div>
</body>
</html>
"@
 
    Send-ResultMail -BodyHtml $bodyHtml
}
catch {
    $errorMessage = $_.Exception.Message
 
    $bodyHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>$emailSubject</title>
</head>
<body style="font-family: Segoe UI, Arial, Helvetica, sans-serif; font-size: 10pt;">
    <h1 style="font-size: 16pt; color: #1f4e79;">Fel vid kontroll av aktiva datorer under Win11 Inaktiva</h1>
    <p><b>Tidpunkt:</b> $(Convert-ToHtmlSafe $runTime)</p>
    <p><b>Sökbas:</b> $(Convert-ToHtmlSafe $searchBase)</p>
    <p><b>Fel:</b> $(Convert-ToHtmlSafe $errorMessage)</p>
</body>
</html>
"@
 
    try {
        Send-ResultMail -BodyHtml $bodyHtml
    }
    catch {
    }
 
    throw
}


