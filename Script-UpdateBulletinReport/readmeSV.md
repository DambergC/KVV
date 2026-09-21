# Update Bulletin Report

Den här mappen innehåller ett schemalagt PowerShell-skript som frågar en SQL-databas efter aktuella uppdateringsbulletiner, filtrerar bort oönskade poster, skapar en HTML-sammanfattning och skickar den via e-post.

## Filer

- `UpdateBulletinReport.ps1` — huvudskriptet som läser konfigurationen, frågar databasen, filtrerar resultatet, skapar HTML och skickar rapporten.
- `UpdateBulletinReport.config.xml` — miljöspecifika inställningar för databasanslutning, filtrering, loggning och e-postleverans.
- `readme.md` — engelsk dokumentation för skriptet och dess konfiguration.
- `readmeSV.md` — svensk dokumentation för skriptet och dess konfiguration.

## Syfte med skriptet

Skriptet är utformat för att:

- ansluta till en SQL Server-instans
- läsa aktuella poster från vyn `v_UpdateInfo`
- filtrera bort utgångna eller ersatta poster
- endast behålla poster inom det konfigurerade datumintervallet (`DaysBack`)
- exkludera bulletinrubriker och beskrivningar som matchar konfigurerade mönster
- skapa en statisk HTML-e-post med en sammanfattningstabell
- skicka resultatet till en eller flera mottagare via SMTP
- logga information om körningen till en loggfil

Detta passar bra för en schemalagd uppgift eller ett Windows Task Scheduler-jobb som skickar ett dagligt eller periodiskt e-postmeddelande med uppdateringsbulletiner.

## Skriptflöde

PowerShell-skriptet följer denna sekvens:

1. Läs XML-konfigurationsfilen.
2. Validera obligatoriska inställningar, till exempel SQL-instans, databasnamn, SMTP-server och mottagare.
3. Läs in nödvändiga moduler.
4. Fråga SQL-databasen med `Invoke-DbaQuery`.
5. Tillämpa exkluderingsfilter på rubriker och beskrivningar.
6. Skapa en temporär HTML-fil i `%TEMP%`.
7. Bygg en formaterad e-posttext med:
   - Publiceringsdatum
   - Artikel-ID
   - Allvarlighetsgrad
   - Rubrik
   - Beskrivning
   - Informations-URL
8. Skicka e-postmeddelandet genom att anropa `Send-MailKitMessage`.
9. Ta bort den temporära HTML-filen om inte `-KeepHtmlFile` används.

## Nödvändiga moduler

Skriptet kontrollerar och importerar uttryckligen två PowerShell-moduler.

### 1) dbatools

Krävs för databasåtkomst och SQL-relaterade hjälpkommandon som:

- `Invoke-DbaQuery`
- `Set-DbatoolsInsecureConnection`

`dbatools` måste vara installerat för det användarkonto som kör den schemalagda uppgiften.

Exempel på installation:

```powershell
Install-Module dbatools -Scope CurrentUser
```

### 2) Send-MailKitMessage

Krävs för att skicka SMTP-e-post med hjälp av MailKit-biblioteket.

Exempel på installation:

```powershell
Install-Module Send-MailKitMessage -Scope CurrentUser
```

Om någon av modulerna saknas avbryts skriptet med ett fel och händelsen loggas.

## Konfigurationsfil

Skriptet läser inställningarna från `UpdateBulletinReport.config.xml`.

Exempel på struktur:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <Database>
    <SqlInstance>servername</SqlInstance>
    <DatabaseName>DB</DatabaseName>
  </Database>

  <Query>
    <DaysBack>1</DaysBack>
    <ExcludeTitleContains>
      <Pattern>25H2</Pattern>
      <Pattern>Defender</Pattern>
    </ExcludeTitleContains>
    <ExcludeDescriptionContains>
      <Pattern>Install the latest version of Windows</Pattern>
    </ExcludeDescriptionContains>
  </Query>

  <Logging>
    <LogPath>d:\Scripts\UpdateBulletinReport\LogFiles\UpdateBulletinReport.log</LogPath>
  </Logging>

  <Email>
    <SmtpServer>smtp.company.se</SmtpServer>
    <SmtpPort>25</SmtpPort>
    <From>noreply@company.se</From>
    <To>
      <Recipient>user@company.se</Recipient>
    </To>
    <Subject>Recent Software Update Bulletins</Subject>
    <HtmlLogoMimeType>image/png</HtmlLogoMimeType>
    <HtmlLogo><![CDATA[data:image/png;base64,...]]></HtmlLogo>
  </Email>
</Configuration>
```

## Beskrivning av inställningarna

### Databasinställningar

#### `Configuration/Database/SqlInstance`

Namnet på SQL Server-instansen eller värdnamnet som används för att ansluta till databasen.

Exempel:

```xml
<SqlInstance>servername</SqlInstance>
```

#### `Configuration/Database/DatabaseName`

Namnet på måldatabasen.

Exempel:

```xml
<DatabaseName>DB</DatabaseName>
```

### Frågeinställningar

#### `Configuration/Query/DaysBack`

Antalet dagar bakåt som ska inkluderas när uppdateringsbulletiner väljs ut.

Exempel:

```xml
<DaysBack>1</DaysBack>
```

Frågan väljer rader där `DatePosted` ligger inom de senaste `N` dagarna.

#### `Configuration/Query/ExcludeTitleContains/Pattern`

Lista över textvärden som gör att en bulletin exkluderas om de förekommer i rubriken.

Exempel på värden i den aktuella konfigurationen:

- `25H2`
- `26H1`
- `arm64`
- `Defender`
- `Apps`
- `Edge-Beta`
- `Edge-Dev`
- `x86`
- `Office 2019`
- `Office LTSC 2021`

Mönstren utvärderas med skiftlägesokänslig delsträngsmatchning.

#### `Configuration/Query/ExcludeDescriptionContains/Pattern`

Lista över textvärden som gör att en bulletin exkluderas om de förekommer i beskrivningen.

Exempel:

```xml
<Pattern>Install the latest version of Windows</Pattern>
```

Detta hjälper till att filtrera bort mindre relevanta eller standardiserade uppdateringsmeddelanden.

### Loggningsinställningar

#### `Configuration/Logging/LogPath`

Sökväg till loggfilen som registrerar skriptets körning och eventuella fel.

Exempel:

```xml
<LogPath>d:\Scripts\UpdateBulletinReport\LogFiles\UpdateBulletinReport.log</LogPath>
```

Loggfilen skapas automatiskt om mappen inte redan finns.

### E-postinställningar

#### `Configuration/Email/SmtpServer`

SMTP-värd som används för att skicka rapporten.

#### `Configuration/Email/SmtpPort`

SMTP-portnummer.

Exempel:

```xml
<SmtpPort>25</SmtpPort>
```

#### `Configuration/Email/From`

Avsändaradress som används i e-posthuvudet.

#### `Configuration/Email/To/Recipient`

En eller flera mottagare. XML-konfigurationen stöder flera mottagarposter.

Exempel:

```xml
<To>
  <Recipient>user1@company.se</Recipient>
  <Recipient>user2@company.se</Recipient>
</To>
```

#### `Configuration/Email/Subject`

Grundämne för e-postmeddelandet. Skriptet lägger automatiskt till datumintervallet, till exempel:

- `Recent Software Update Bulletins - 2026-09-01 to 2026-09-15`

#### `Configuration/Email/HtmlLogoMimeType`

MIME-typ för den inbäddade logotypen.

Exempel på MIME-typer som stöds:

- `image/png`
- `image/jpeg`
- `image/gif`

#### `Configuration/Email/HtmlLogo`

Logotyp som används i HTML-e-postmeddelandet. Värdet kan vara:

- en komplett data-URI (`data:image/png;base64,...`)
- en rå base64-sträng
- en HTTP- eller HTTPS-URL

Skriptet skapar en HTML-`<img>`-tagg med den konfigurerade logotypen när värdet är giltigt.

## Hantering av allvarlighetsgrad

Skriptet översätter databasens värden för allvarlighetsgrad till läsbara etiketter i HTML-rapporten:

- `10` = Kritisk
- `8` = Viktig
- `6` = Måttlig
- `2` = Låg
- `0` = Ospecificerad

Rapporten innehåller även text om CVSS-intervall samt en kort beskrivning för varje allvarlighetsgrad.

## HTML-utdata

Det genererade e-postmeddelandet innehåller en tabell med följande kolumner:

- Publiceringsdatum
- Artikel-ID
- Allvarlighetsgrad
- Rubrik
- Beskrivning
- Informations-URL

Om inga bulletiner hittas inom det valda datumintervallet innehåller e-postmeddelandet ett enkelt meddelande:

> Inga nya bulletiner hittades under de senaste X dagarna.

## Skriptparametrar

Skriptet stöder följande parametrar:

```powershell
UpdateBulletinReport.ps1 [-ConfigPath <string>] [-KeepHtmlFile]
```

### `-ConfigPath`

Valfri sökväg till XML-konfigurationsfilen. Standardvärde:

```powershell
D:\Scripts\UpdateBulletinReport\UpdateBulletinReport.config.xml
```

### `-KeepHtmlFile`

Om parametern anges tas den genererade temporära HTML-filen inte bort efter att e-postmeddelandet har skickats.

## Driftsinformation

- Skriptet skriver detaljerade loggar till den konfigurerade sökvägen.
- Det är utformat för att köras som en schemalagd uppgift utan krav på en interaktiv konsol.
- Saknade obligatoriska konfigurationsvärden behandlas som kritiska fel.
- Skriptet validerar även att SMTP-porten är numerisk och ligger inom ett giltigt intervall.
- Skriptet använder `Set-StrictMode -Version Latest` och `$ErrorActionPreference = "Stop"` för striktare körningssäkerhet.

## Exempel på användning

Kör från PowerShell:

```powershell
.\UpdateBulletinReport.ps1
```

Eller med en anpassad konfigurationsfil:

```powershell
.\UpdateBulletinReport.ps1 -ConfigPath "D:\Scripts\UpdateBulletinReport\Custom.config.xml"
```

För att behålla den genererade HTML-filen:

```powershell
.\UpdateBulletinReport.ps1 -KeepHtmlFile
```

## Sammanfattning

Detta skript är ett lättviktigt verktyg för schemalagd rapportering av Microsofts uppdateringsbulletiner. Det samlar in relevanta bulletinposter, filtrerar dem enligt konfigurerade regler, formaterar resultatet som ett professionellt HTML-e-postmeddelande och skickar det till angivna mottagare.

De viktigaste värdena som normalt behöver anpassas är:

- SQL-server och databas
- rapporteringsintervall
- exkluderingsnyckelord
- loggsökväg
- SMTP-inställningar
- mottagare
- valfri företagslogotyp

Alla dessa inställningar finns samlade i XML-konfigurationsfilen.
