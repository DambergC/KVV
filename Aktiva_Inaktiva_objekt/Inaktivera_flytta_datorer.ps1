Import-Module ActiveDirectory -ErrorAction Stop

 

$ErrorActionPreference = 'Stop'

 

# =========================================

# Konfiguration

# Inaktivera-flytta-datorobjekt

# =========================================

 

$simulationMode = $false

 

$searchBase = "OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

 

$disableAfterDays = 50

$moveAfterDays    = 90

 

$disableDateFilter = (Get-Date).AddDays(-$disableAfterDays)

$moveDateFilter    = (Get-Date).AddDays(-$moveAfterDays)

 

# Generell mål-OU för allt som inte matchar specialregler direkt

$defaultTargetOU = "OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

 

# Specialregler gäller ENDAST om datorobjektet ligger direkt i SourceOU

$ouMoveMap = @(

    [pscustomobject]@{

        SourceOU = "OU=DMITA,OU=MITA,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

        TargetOU = "OU=DeladMita,OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

    },

    [pscustomobject]@{

        SourceOU = "OU=MITA,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

        TargetOU = "OU=Mita,OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

    },

    [pscustomobject]@{

        SourceOU = "OU=T1,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

        TargetOU = "OU=T1,OU=Win11,OU=Inaktiva,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

    }

)

 

$excludedOUs = @(

    "OU=DP,OU=T1,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se",

    "OU=DP-Server,OU=T1,OU=Fysiska klienter,OU=Windows 11,OU=Datorer,OU=Kriminalvården,DC=kvv,DC=se"

)

 

$emailRecipients = "HK.IT.CMKalendern@kriminalvarden.se"

$emailSubject    = "AD - Hantering av inaktiva Windows 11-datorobjekt"

$mailFrom        = "noreply@kriminalvarden.se"

$smtpServer      = "smtp.kvv.se"

 

$maxObjectsPerRun = 4000

 

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

 

function Get-TargetOUFromDN {

    param(

        [Parameter(Mandatory = $true)]

        [string]$DistinguishedName,

 

        [Parameter(Mandatory = $true)]

        [array]$MoveMap,

 

        [Parameter(Mandatory = $true)]

        [string]$DefaultTargetOU

    )

 

    $parentOU = Get-ParentOUFromDN -DistinguishedName $DistinguishedName

 

    foreach ($mapping in $MoveMap) {

        if ($parentOU -ieq $mapping.SourceOU) {

            return $mapping.TargetOU

        }

    }

 

    return $DefaultTargetOU

}

 

function Test-IsInExcludedOU {

    param(

        [Parameter(Mandatory = $true)]

        [string]$DistinguishedName,

 

        [Parameter(Mandatory = $true)]

        [string[]]$ExcludedOUs

    )

 

    foreach ($excludedOU in $ExcludedOUs) {

        if ($DistinguishedName -like "*,$excludedOU") {

            return $true

        }

    }

 

    return $false

}

 

function Send-ResultMail {

    param(

        [Parameter(Mandatory = $true)]

        [string]$Subject,

 

        [Parameter(Mandatory = $true)]

        [string]$BodyHtml

    )

 

    Send-MailMessage `

        -From $mailFrom `

        -To $emailRecipients `

        -Subject $Subject `

        -Encoding UTF8 `

        -SmtpServer $smtpServer `

        -BodyAsHtml `

        -Body $BodyHtml

}

 

function New-ResultMailBody {

    param(

        [Parameter(Mandatory = $true)]

        [string]$RunTime,

 

        [Parameter(Mandatory = $true)]

        [string]$SearchBase,

 

        [Parameter(Mandatory = $true)]

        [string]$DefaultTargetOU,

 

        [Parameter(Mandatory = $true)]

        [string[]]$ExcludedOUs,

 

        [Parameter(Mandatory = $true)]

        [int]$DisableAfterDays,

 

        [Parameter(Mandatory = $true)]

        [int]$MoveAfterDays,

 

        [Parameter(Mandatory = $true)]

        [int]$RetrievedCount,

 

        [Parameter(Mandatory = $true)]

        [int]$DisabledNowCount,

 

        [Parameter(Mandatory = $true)]

        [int]$AlreadyDisabledCount,

 

        [Parameter(Mandatory = $true)]

        [int]$MovedNowCount,

 

        [Parameter(Mandatory = $true)]

        [int]$FailedItems,

 

        [Parameter(Mandatory = $true)]

        [array]$ProcessedResults,

 

        [string]$InfoMessage = ""

    )

 

    $excludedOUsHtml = ($ExcludedOUs | ForEach-Object {

        Convert-ToHtmlSafe $_

    }) -join "<br />"

 

    $tableRows = foreach ($item in $ProcessedResults) {

@"

<tr>

    <td>$(Convert-ToHtmlSafe $item.Name)</td>

    <td>$(Convert-ToHtmlSafe $item.DNSHostName)</td>

    <td>$(Convert-ToHtmlSafe $item.SamAccountName)</td>

    <td>$(Convert-ToHtmlSafe $item.Enabled)</td>

    <td>$(Convert-ToHtmlSafe $item.LastLogonDate)</td>

    <td>$(Convert-ToHtmlSafe $item.DisableStatus)</td>

    <td>$(Convert-ToHtmlSafe $item.MoveStatus)</td>

    <td>$(Convert-ToHtmlSafe $item.ResultMessage)</td>

</tr>

"@

    }

 

    if (-not $tableRows) {

        $tableRows = @"

<tr>

    <td colspan="8">Inga datorobjekt hittades som varit utan logon i mer än $DisableAfterDays dagar.</td>

</tr>

"@

    }

 

    $infoSection = ""

    if (-not [string]::IsNullOrWhiteSpace($InfoMessage)) {

        $infoSection = "<p><b>Information:</b> $(Convert-ToHtmlSafe $InfoMessage)</p>"

    }

 

    return @"

<html>

<head>

<meta charset="UTF-8">

<style>

body { font-family: Segoe UI, Arial, sans-serif; font-size: 10pt; }

table { border-collapse: collapse; width: 100%; }

th, td { border: 1px solid #d0d0d0; padding: 6px; text-align: left; vertical-align: top; }

th { background-color: #f0f0f0; }

</style>

</head>

<body>

    <h3>Resultat från skript: Inaktiva Windows 11-datorer</h3>

    <p>

        <b>Körtid:</b> $(Convert-ToHtmlSafe $RunTime)<br />

        <b>Sökbas:</b> $(Convert-ToHtmlSafe $SearchBase)<br />

        <b>Mål-OU:</b> $(Convert-ToHtmlSafe $DefaultTargetOU)<br />

        <b>Exkluderade OU:er:</b><br />

        $excludedOUsHtml

    </p>

 

    $infoSection

 

    <p>

        <b>Inaktiveringsgräns:</b> $DisableAfterDays dagar<br />

        <b>Flyttgräns:</b> $MoveAfterDays dagar<br />

        <b>Antal träffar:</b> $RetrievedCount<br />

        <b>Nyinaktiverade:</b> $DisabledNowCount<br />

        <b>Redan inaktiverade:</b> $AlreadyDisabledCount<br />

        <b>Nyflyttade:</b> $MovedNowCount<br />

        <b>Fel:</b> $FailedItems

    </p>

 

    <table>

        <tr>

            <th>Name</th>

            <th>DNSHostName</th>

            <th>SamAccountName</th>

            <th>Enabled</th>

            <th>LastLogonDate</th>

            <th>DisableStatus</th>

            <th>MoveStatus</th>

            <th>ResultMessage</th>

        </tr>

        $($tableRows -join "`r`n")

    </table>

</body>

</html>

"@

}

 

# =========================================

# Huvudlogik

# =========================================

 

$runTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$processedResults = @()

 

try {

    # Validera mål-OU:er

    Get-ADOrganizationalUnit -Identity $defaultTargetOU -ErrorAction Stop | Out-Null

 

    foreach ($mapping in $ouMoveMap) {

        Get-ADOrganizationalUnit -Identity $mapping.TargetOU -ErrorAction Stop | Out-Null

    }

 

    # Hämta datorobjekt

    $computersToProcess = Get-ADComputer `

        -Filter { LastLogonDate -lt $disableDateFilter } `

        -SearchBase $searchBase `

        -SearchScope Subtree `

        -Properties Name, DNSHostName, Enabled, LastLogonDate, sAMAccountName, DistinguishedName |

        Where-Object {

            -not (Test-IsInExcludedOU -DistinguishedName $_.DistinguishedName -ExcludedOUs $excludedOUs)

        } |

        Select-Object Name, DNSHostName, Enabled, LastLogonDate, sAMAccountName, DistinguishedName

 

    $computerCount = @($computersToProcess).Count

 

    if ($computerCount -gt $maxObjectsPerRun) {

        $bodyHtml = New-ResultMailBody `

            -RunTime $runTime `

            -SearchBase $searchBase `

            -DefaultTargetOU $defaultTargetOU `

            -ExcludedOUs $excludedOUs `

            -DisableAfterDays $disableAfterDays `

            -MoveAfterDays $moveAfterDays `

            -RetrievedCount $computerCount `

            -DisabledNowCount 0 `

            -AlreadyDisabledCount 0 `

            -MovedNowCount 0 `

            -FailedItems 1 `

            -ProcessedResults @() `

            -InfoMessage "Körningen avbröts eftersom antalet träffar ($computerCount) översteg maxgränsen ($maxObjectsPerRun). Inga ändringar gjordes."

 

        Send-ResultMail -Subject $emailSubject -BodyHtml $bodyHtml

        exit 1

    }

 

    foreach ($computer in $computersToProcess) {

        $disableStatus    = "NotRequired"

        $moveStatus       = "NotRequired"

        $resultMessage    = ""

        $shouldMove       = $false

        $isInactiveForMove = $false

        $resolvedTargetOU = $null

 

        if ($computer.LastLogonDate -lt $moveDateFilter) {

            $shouldMove = $true

            $resolvedTargetOU = Get-TargetOUFromDN `

                -DistinguishedName $computer.DistinguishedName `

                -MoveMap $ouMoveMap `

                -DefaultTargetOU $defaultTargetOU

        }

 

        if (-not $simulationMode) {

            if ($computer.Enabled) {

                try {

                    Disable-ADAccount -Identity $computer.DistinguishedName -ErrorAction Stop

                    $disableStatus = "Success"

                }

                catch {

                    $disableStatus = "Failed"

                    $resultMessage = "Disable failed: $($_.Exception.Message)"

                }

            }

            else {

                $disableStatus = "AlreadyDisabled"

            }

 

            if ($disableStatus -in @("Success", "AlreadyDisabled")) {

                $isInactiveForMove = $true

            }

 

            if ($shouldMove) {

                if ($isInactiveForMove) {

                    if ($computer.DistinguishedName -like "*,$resolvedTargetOU") {

                        $moveStatus = "AlreadyInTargetOU"

                    }

                    else {

                        try {

                            Move-ADObject -Identity $computer.DistinguishedName -TargetPath $resolvedTargetOU -ErrorAction Stop

                            $moveStatus = "Success"

                        }

                        catch {

                            $moveStatus = "Failed"

                            if ([string]::IsNullOrWhiteSpace($resultMessage)) {

                                $resultMessage = "Move failed to '$resolvedTargetOU': $($_.Exception.Message)"

                            }

                            else {

                                $resultMessage += " | Move failed to '$resolvedTargetOU': $($_.Exception.Message)"

                            }

                        }

                    }

                }

                else {

                    $moveStatus = "SkippedBecauseComputerIsNotInactive"

                }

            }

        }

        else {

            if ($computer.Enabled) {

                $disableStatus = "WouldDisable"

                $isInactiveForMove = $true

            }

            else {

                $disableStatus = "AlreadyDisabled"

                $isInactiveForMove = $true

            }

 

            if ($shouldMove) {

                if ($isInactiveForMove) {

                    if ($computer.DistinguishedName -like "*,$resolvedTargetOU") {

                        $moveStatus = "AlreadyInTargetOU"

                    }

                    else {

                        $moveStatus = "WouldMove"

                    }

                }

                else {

                    $moveStatus = "SkippedBecauseComputerIsNotInactive"

                }

            }

            else {

                $moveStatus = "NotRequired"

            }

 

            $resultMessage = "Simulation mode - no changes made"

        }

 

        $processedResults += [pscustomobject]@{

            Name           = $computer.Name

            DNSHostName    = $computer.DNSHostName

            SamAccountName = $computer.sAMAccountName

            Enabled        = $computer.Enabled

            LastLogonDate  = $computer.LastLogonDate

            DisableStatus  = $disableStatus

            MoveStatus     = $moveStatus

            ResultMessage  = $resultMessage

        }

    }

 

    $retrievedCount       = @($processedResults).Count

    $disabledNowCount     = @($processedResults | Where-Object { $_.DisableStatus -eq "Success" }).Count

    $alreadyDisabledCount = @($processedResults | Where-Object { $_.DisableStatus -eq "AlreadyDisabled" }).Count

    $movedNowCount        = @($processedResults | Where-Object { $_.MoveStatus -eq "Success" }).Count

    $failedItems          = @($processedResults | Where-Object { $_.DisableStatus -eq "Failed" -or $_.MoveStatus -eq "Failed" }).Count

 

    $bodyHtml = New-ResultMailBody `

        -RunTime $runTime `

        -SearchBase $searchBase `

        -DefaultTargetOU $defaultTargetOU `

        -ExcludedOUs $excludedOUs `

        -DisableAfterDays $disableAfterDays `

        -MoveAfterDays $moveAfterDays `

        -RetrievedCount $retrievedCount `

        -DisabledNowCount $disabledNowCount `

        -AlreadyDisabledCount $alreadyDisabledCount `

        -MovedNowCount $movedNowCount `

        -FailedItems $failedItems `

        -ProcessedResults $processedResults

 

    Send-ResultMail -Subject $emailSubject -BodyHtml $bodyHtml

    exit 0

}

catch {

    $errorMessage = $_.Exception.Message

 

    $bodyHtml = New-ResultMailBody `

        -RunTime $runTime `

        -SearchBase $searchBase `

        -DefaultTargetOU $defaultTargetOU `

        -ExcludedOUs $excludedOUs `

        -DisableAfterDays $disableAfterDays `

        -MoveAfterDays $moveAfterDays `

        -RetrievedCount 0 `

        -DisabledNowCount 0 `

        -AlreadyDisabledCount 0 `

        -MovedNowCount 0 `

        -FailedItems 1 `

        -ProcessedResults @() `

        -InfoMessage "Skriptet avslutades med fel: $errorMessage"

 

    try {

        Send-ResultMail -Subject $emailSubject -BodyHtml $bodyHtml

    }

    catch {

    }

 

    exit 1

}