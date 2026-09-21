# Server Compliance Report

Detta PowerShell-skript kontrollerar om servrar i en eller flera Configuration Manager-samlingar är kompatibla med en angiven Software Update Group (SUG) och skickar en sammanfattad HTML-e-postrapport.

Skriptet är avsett att köras som en schemalagd uppgift dagligen. Det utvärderar serverkompatibilitet i förhållande till den aktuella eller senaste Patch Tuesday och skickar endast e-post för samlingar som är konfigurerade för den aktuella dagen.

## Filer

- `ServerComplianceReport.ps1` – huvudrapportsskriptet
- `ServerComplianceReport.config.xml` – XML-konfiguration för databas, samlingar, loggning och e-postmottagare
- `Readme.md` – dokumentation för skriptet

## Vad skriptet gör

- Läser XML-konfigurationsfilen
- Bestämmer aktuellt Patch Tuesday och beräknar den konfigurerade offseten från detta datum
- Frågar ConfigMgr-databasen efter samlingsmedlemmar och kompatibilitetsstatus
- Identifierar:
  - kompatibla servrar
  - icke-kompatibla servrar
  - servrar med okänd kompatibilitetsstatus
- Bygger en HTML-rapport per samling
- Skickar e-post för varje matchande samling via SMTP
- Behåller valfritt genererade HTML-filer i `%TEMP%` när du använder flaggan `-KeepHtmlFile`

## Krav

Skriptet kräver att följande PowerShell-moduler är installerade för kontot som kör den schemalagda uppgiften:

```powershell
Install-Module dbatools -Scope CurrentUser
Install-Module Send-MailKitMessage -Scope CurrentUser
```

Skriptet förutsätter även åtkomst till SQL-databasen som används av Configuration Manager och en giltig SMTP-server.

## Konfiguration

Skriptet använder XML-filen `ServerComplianceReport.config.xml`.

Exempelstruktur:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <Database>
    <SqlInstance>servername</SqlInstance>
    <DatabaseName>DB</DatabaseName>
  </Database>

  <Query>
    <Collections>
      <Collection>
        <Name>SU - Server Updates - Server 100</Name>
        <Title>Server Patch Tuesday</Title>
        <XDays>10</XDays>
        <DaysAfterPatchTuesday>6</DaysAfterPatchTuesday>
      </Collection>
    </Collections>
  </Query>

  <Logging>
    <LogPath>d:\Scripts\ServerComplianceReport\LogFiles\ServerComplianceReport.log</LogPath>
  </Logging>

  <Email>
    <SmtpServer>smtp.company.se</SmtpServer>
    <SmtpPort>25</SmtpPort>
    <From>noreply@company.se</From>
    <To>
      <Recipient>user@company.se</Recipient>
    </To>
    <Subject>Server Patch Compliance Report</Subject>
    <HtmlLogoMimeType>image/png</HtmlLogoMimeType>
    <HtmlLogo><![CDATA[data:image/png;base64,...]]></HtmlLogo>
  </Email>
</Configuration>
```

### Databasinställningar

- `SqlInstance` – SQL Server/instansnamn som används av dbatools
- `DatabaseName` – databas som innehåller ConfigMgr-data

### Samlingsinställningar

Varje `<Collection>`-post definierar en samling och rapporteringsschema:

- `Name` – namn på ConfigMgr-samlingen
- `Title` – namn på Software Update Group som ska kontrolleras
- `XDays` – maximalt antal dagar en server får vara "hälsosam/online" för att anses frisk
- `DaysAfterPatchTuesday` – hur många dagar efter Patch Tuesday som denna samling ska köras

Skriptet jämför aktuellt datum mot den konfigurerade offseten och kör endast de samlingar som matchar det värdet.

### Loggningsinställningar

- `LogPath` – sökväg till loggfilen som används av skriptet

Om katalogen saknas skapar skriptet den automatiskt.

### E-postinställningar

- `SmtpServer` – SMTP-serverns värdnamn
- `SmtpPort` – SMTP-portnummer
- `From` – avsändaradress för e-post
- `To/Recipient` – en eller flera mottagare
- `Subject` – ämnesprefix för genererade e-postmeddelanden
- `HtmlLogo` – valfri Base64-bild eller URL för ett e-postlogotyp
- `HtmlLogoMimeType` – MIME-typ för den inbäddade logotypen

## Obligatoriska moduler

Skriptet importerar och validerar följande moduler innan det körs:

- `dbatools`
- `Send-MailKitMessage`

Skriptet kontrollerar dessa moduler i kontexten för den schemalagda uppgiftens användarkonto och kastar ett fel om de saknas.

## Exekvering

Kör skriptet från PowerShell:

```powershell
.\ServerComplianceReport.ps1
```

Valfria parametrar:

```powershell
.\ServerComplianceReport.ps1 -ConfigPath "D:\Scripts\ServerComplianceReport\ServerComplianceReport.config.xml"
.\ServerComplianceReport.ps1 -KeepHtmlFile
.\ServerComplianceReport.ps1 -AsOfDate (Get-Date)
```

## Anmärkningar

- Skriptet är utformat för att användas med schemalagda aktiviteter.
- Om ingen konfigurerad samling matchar dagens Patch Tuesday-offset avslutar skriptet utan att skicka e-post.
- Frågeresultat visas i en sammanfattningstabell med:
  - totalt utvärderade servrar
  - kompatibla servrar
  - icke-kompatibla servrar
  - servrar med okänd kompatibilitetsstatus
- E-postrapporter skickas en per matchande samling.

## Exempel på användningsfall

En typisk miljö kan definiera flera serversamlingar med olika Patch Tuesday-offsetvärden. Till exempel:

- Samling A körs 6 dagar efter Patch Tuesday
- Samling B körs 8 dagar efter Patch Tuesday
- Samling C körs 15 dagar efter Patch Tuesday

Detta gör att en rapport kan skickas vid rätt tidpunkt för varje servergrupp utan att skriptet måste köras manuellt varje dag.

## Felsökning

Vanliga problem:

- Modul saknas för användaren som kör den schemalagda uppgiften
- XML-fil saknas eller är felaktigt formaterad
- SMTP-servern kan inte nås
- Ingen samling matchar det beräknade offsetvärdet för dagen
- ConfigMgr-samling eller SUG-namn matchar inte de konfigurerade värdena

Skriptet skriver detaljerad logg till den konfigurerade loggvägen och använder `Write-Error`/`Write-Warning` när problem uppstår.
