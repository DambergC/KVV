# Send-WindowsUpdateDeployed

Ett PowerShell-automationskript för rapportering av Windows Update-distributioner via Microsoft Endpoint Configuration Manager (MECM/SCCM). Skriptet genererar och skickar HTML-rapporter via e-post över distribuerade uppdateringar.

## Översikt

Skriptet läser sina inställningar från en XML-konfigurationsfil och:

- Frågar MECM efter programuppdateringar i en angiven Software Update Group
- Bygger en formaterad HTML-rapport (PSWriteHTML)
- Skickar rapporten via SMTP (Send-MailKitMessage)

Det är utformat för att köras på en MECM-site server/konsole-maskin, antingen manuellt eller som en Schemalagd aktivitet (Scheduled Task).

## Nya / uppdaterade funktioner (Skript + XML)

### Konfigurationsfil v2 (krävs)

PowerShell-skriptet förväntar sig för närvarande en XML-fil med namnet:

- `Send-WindowsUpdateDeployedv2.xml` (måste ligga i samma mapp som `Send-WindowsUpdateDeployed.ps1`)

> Obs: Repositoriet innehåller även `Send-WindowsUpdateDeployed.XML` som exempel/mall, men skriptet är hårdkodat att läsa in `Send-WindowsUpdateDeployedv2.xml`.

### Hjälpfunktioner för XML-tolkning (PowerShell 5.1-kompatibelt)

Skriptet innehåller små hjälpfunktioner för att pålitligt läsa XML-värden:

- `Get-XmlText` (trimmade strängvärden)
- `Get-XmlInt` (heltal med standardvärden)
- `Get-XmlIntArray` (lista av heltal)

Det gör konfigurationen mer robust mot saknade/tomma värden.

### Kompatibilitet för mottagarschema (nytt)

Skriptet stöder **två** mottagarscheman:

1) **Nästlat element**-format:

```xml
<Recipients>
  <Recipients>
    <Email>recipient@company.com</Email>
  </Recipients>
</Recipients>
```

2) **Attribut**-format (används av det inkluderade XML-exemplet):

```xml
<Recipients>
  <Recipients email="recipient@company.com" />
</Recipients>
```

Mottagare avdupliceras automatiskt.

### Kompatibilitet för DisableReportMonth-schema (nytt)

Skriptet stöder **två** sätt att ange inaktiverade månader:

1) Elementformat:

```xml
<DisableReportMonth>
  <DisableReportMonth>
    <Number>7</Number>
  </DisableReportMonth>
</DisableReportMonth>
```

2) Attributformat (används av det inkluderade XML-exemplet):

```xml
<DisableReportMonth>
  <DisableReportMonth Number="7" />
</DisableReportMonth>
```

Om aktuell månad är inaktiverad avslutas skriptet tidigt.

### Inbyggd validering av konfiguration (nytt)

Skriptet validerar obligatoriska inställningar vid start och avbryter tidigt med tydliga felmeddelanden om något saknas:

- `SiteServer`
- `MailSMTP`
- `Mailfrom`
- `MailCustomer`
- `UpdateDeployed/UpdateGroupName`
- Minst en mottagare

### Valfri loggning (nytt)

Loggning är valfritt. Om `Logfile/Path` och `Logfile/Name` finns angivna skriver skriptet tidsstämplade loggrader.

- Loggformat: `yyyy/MM/dd HH:mm:ss <meddelande>`

### Automatisk kontroll + import av moduler (nytt)

Skriptet säkerställer att nödvändiga moduler är installerade och importerar dem:

- `PSWriteHTML`
- `Send-MailKitMessage`

Om en modul saknas stoppar skriptet med ett tydligt meddelande.

### Automatisk identifiering av MECM CMSite PSDrive (nytt)

I stället för att kräva en hårdkodad site-kod upptäcker skriptet automatiskt den första `CMSite`-PSDriven och växlar temporärt till den för att köra:

- `Get-CMSoftwareUpdate -Fast -UpdateGroupName <name>`

Det förbättrar portabiliteten mellan MECM-miljöer.

### Uppdaterad logik för datainsamling (nytt)

Skriptet bygger rapportens dataset genom att välja vanliga uppdateringsegenskaper från MECM:

- `ArticleID`
- `LocalizedDisplayName` (visas som `Title`)
- `LocalizedDescription`
- `DatePosted`
- `IsDeployed` (visas som `Deployed`)
- `LocalizedInformativeURL` (visas som `URL`)
- `SeverityName` (visas som `Severity`)

### HTML-rapport omgjord för att matcha DPStatus-stil (nytt)

Rapporten genereras med **PSWriteHTML** och innehåller:

- Inbäddad base64-bild/logotyp
- En "Sammanfattning"-sektion som visar:
  - Datum (`yyyy-MM-dd`)
  - SiteServer
  - UpdateGroupName
  - Antal uppdateringar som returnerats
  - `LimitDays`
  - `DaysAfterPatchTuesdayToReport`
- En tabell med uppdateringar (eller "Inga poster." om tom)
- En sidfot med skapandetidsstämpel och datornamn

### E-postutskick uppdaterat till Send-MailKitMessage + MimeKit (nytt)

Skriptet bygger en `MimeKit.InternetAddressList` för mottagare och skickar HTML-innehållet med `Send-MailKitMessage`.

Ämnesradens format:

- `<MailCustomer> - Windows Updates <Månadsnamn> <År> - <yyyy-MM-dd>`

## Förutsättningar

### Nödvändiga moduler

Installera modulerna (på maskinen som kör den schemalagda aktiviteten/skriptet):

```powershell
Install-Module PSWriteHTML
Install-Module send-mailkitmessage
```

### Systemkrav

- Windows Server med MECM Site Server eller Console installerad
- PowerShell 5.1 eller senare
- Nätverksåtkomst till MECM site server
- Åtkomst till SMTP-server för att skicka e-post
- Lämpliga MECM-behörigheter för att fråga efter uppdateringsinformation

## Installation

1. Klona/ladda ned mappen till MECM site server / konsole-maskinen.
2. Skapa `Send-WindowsUpdateDeployedv2.xml` i samma katalog som `Send-WindowsUpdateDeployed.ps1`.
3. Installera de nödvändiga PowerShell-modulerna.
4. Konfigurera XML-filen.
5. (Valfritt) Skapa en schemalagd aktivitet för att köra skriptet.

## XML-konfigurationsfil

Repositoriet innehåller `Send-WindowsUpdateDeployed.XML` som exempel. Skriptet förväntar sig `Send-WindowsUpdateDeployedv2.xml`.

### Exempel-XML (attributbaserade mottagare + inaktiverade månader)

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
	<Logfile>
		<Path>G:\Scripts\Logfiles\</Path>
		<Name>WindowsUpdateScript.log</Name>
		<Logfilethreshold>2000000</Logfilethreshold>
	</Logfile>
	<HTMLfilePath>G:\Scripts\OutFiles\</HTMLfilePath>
	<RunScript>
		<Job DeploymentID="16777362" Offsetdays="2" Description="Grupp100"/>
		<Job DeploymentID="16777363" Offsetdays="8" Description="Grupp200"/>
		<Job DeploymentID="16777364" Offsetdays="15" Description="Grupp300"/>
	</RunScript>
	<DisableReportMonth>
		<DisableReportMonth Number=""/>
		<DisableReportMonth Number=""/>
		<DisableReportMonth Number=""/>
	</DisableReportMonth>
	<Recipients>
		<Recipients email="christian.damberg@kriminalvarden.se"/>
	</Recipients>
	<UpdateDeployed>
		<LimitDays>-25</LimitDays>
		<DaysAfterPatchToRun>9</DaysAfterPatchToRun>
		  <UpdateGroups>
    <UpdateGroupName>Server Patch Tuesday</UpdateGroupName>
    <UpdateGroupName>Windows 11 OS Patch</UpdateGroupName>
	<UpdateGroupName>Windows 11 Office Patch</UpdateGroupName>
  </UpdateGroups>
	</UpdateDeployed>
	<SiteServer>Vntapp0780</SiteServer>
	<Mailfrom>no-reply@kvv.se</Mailfrom>
	<MailSMTP>smtp.kvv.se</MailSMTP>
	<MailPort>25</MailPort>
	<MailCustomer>Kriminalvarden - IT</MailCustomer>
</Configuration>
```

## Användning

### Manuell körning

```powershell
cd C:\Path\To\Script-WindowsUpdateDeployed
.\Send-WindowsUpdateDeployed.ps1
```

### Schemalagd aktivitet (Scheduled Task)

- Program: `PowerShell.exe`
- Argument: `-ExecutionPolicy Bypass -File "C:\Path\To\Script-WindowsUpdateDeployed\Send-WindowsUpdateDeployed.ps1"`
- Start i: `C:\Path\To\Script-WindowsUpdateDeployed`
- Kör med högsta behörighet: Aktiverad

## Felsökning

### Get-CMSoftwareUpdate är inte tillgängligt

- Säkerställ att MECM-konsolen är installerad
- Importera modulen `ConfigurationManager` och bekräfta att en `CMSite`-PSDrive finns

### Inga mottagare hittades

Säkerställ att du använder antingen:

- `/Configuration/Recipients/Recipients/Email`-noder, eller
- `/Configuration/Recipients/Recipients[@email]`-attribut
