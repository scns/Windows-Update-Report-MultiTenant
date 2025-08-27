# Windows Update Report MultiTenant

| Repository Status | Windows Update Report |
| :--- | :--- |
|  [![last commit time][github-last-commit]][github-master] [![GitHub Activity][commits-shield]][commits] | |
| [![License][license-shield]](LICENSE) [![Forks][forks-shield]][forks-url] [![Stargazers][stars-shield]][stars-url] [![Issues][issues-shield]][issues-url] | [![Contributors][contributors-shield]][contributors-url] [![GitHub release](https://img.shields.io/github/release/scns/Windows-Update-Report-MultiTenant.svg)](https://GitHub.com/scns/Windows-Update-Report-MultiTenant/releases) |

![Dashboard voorbeeld](images/001.png)

Dit PowerShell-project genereert een uitgebreid overzichtsrapport van Windows Update en device compliance status voor meerdere tenants via Microsoft Graph. Het resultaat is een professioneel HTML-dashboard met filterbare tabellen, grafieken, compliance monitoring en gedetailleerde KB informatie.

## 🚀 Hoofdfunctionaliteit

- **Automatische module installatie**: Controleert en installeert automatisch benodigde PowerShell modules
- **Configureerbare instellingen**: Alle instellingen beheerbaar via `config.json`
- **Multi-tenant ondersteuning**: Haalt per tenant Windows Update en compliance status op via Microsoft Graph
- **Device Compliance Monitoring**: Complete device compliance status tracking via Microsoft Graph API
- **Intelligente KB detectie**: Toont specifieke ontbrekende KB nummers en security patches
- **KB Mapping Database**: Uitgebreide online database met intelligent caching systeem
- **OS versie analyse**: Automatische detectie van verouderde Windows builds en aanbevelingen
- **Timezone ondersteuning**: Configureerbare tijdzone conversie voor accurate LastSeen tijden
- **Flexibele export opties**: Exporteert resultaten naar CSV-bestanden per klant inclusief compliance data
- **Interactief HTML-dashboard**: Professioneel dashboard met filterbare tabellen, snelfilters en grafieken
- **Intelligente bestandsbeheer**: Automatische archivering van oude export bestanden
- **Automatische browser integratie**: Configureerbaar automatisch openen van het gegenereerde rapport

## 🆕 Nieuwe Functionaliteiten v3.0

### 🛡️ **Device Compliance Monitoring**

- **Microsoft Graph Integration**: Volledige integratie met `deviceCompliancePolicyStates` API
- **Compliance Status Tracking**:
  - **Compliant**: Device voldoet aan alle compliance policies
  - **Non-Compliant**: Device heeft compliance issues gedetecteerd
  - **Geen data**: Geen compliance informatie beschikbaar
  - **Error**: Fout opgetreden tijdens compliance controle
- **Visual Indicators**: Kleurgecodeerde compliance status (groen/rood/grijs/oranje)
- **Dedicated Filtering**: Non-Compliant quick filter voor snelle problem identification
- **CSV Export**: Compliance status opgenomen in alle export bestanden
- **Dropdown Filters**: Compliance Status kolom heeft eigen dropdown filter

### 🕐 **Timezone Support**

- **Configureerbare Offset**: Instelbare tijdzone via `timezoneOffsetHours` in config.json
- **Robuuste Conversie**: Ondersteunt meerdere DateTime formaten voor maximale compatibiliteit
- **UTC Detection**: Intelligente tijdzone detectie en conversie
- **Visual Feedback**: HTML headers tonen tijdzone informatie (bijv. "LastSeen (UTC+2)")
- **Accurate Calculations**: Verbeterde sync berekeningen met tijdzone correctie

### 🎯 **Intelligente Update Detectie**

- **Specifieke KB nummers**: Toont ontbrekende KB updates zoals "KB5041585" voor machines met verouderde OS
- **Build analyse**: Analyseert OS versie verschillen en suggereert benodigde cumulative updates
- **Update status categorieën**:
  - "Up-to-date", "Verouderde OS versie", "Handmatige controle vereist"
  - "Waarschijnlijk up-to-date", "Updates wachtend", "Update fouten"
  - "Compliance problemen", "Synchronisatie vereist", "Error"

### 🗄️ **KB Mapping Database & Intelligent Caching**

- **Online KB database**: Uitgebreide mapping van Windows build numbers naar specifieke KB updates
- **Intelligent caching systeem**: Downloads KB database eenmalig en cached voor configureerbare duur (standaard 30 minuten)
- **Fallback mechanisme**: Gebruikt expired cache bij netwerk problemen voor betrouwbaarheid
- **Database overzicht**: Dedicated dashboard tab toont beschikbare mappings en cache status
- **Performance optimalisatie**: Vermindert server load met 95%+ door slim caching
- **Multi-platform support**: Ondersteunt Windows 10, Windows 11 en historische versies
- **Cache methode tracking**: Toont bron van KB informatie (Online, Cache, ExpiredCache, Local, Estimated)

### 🔍 **Geavanceerde Filtering & UI**

- **Dropdown filters**: Update Status en Compliance Status kolommen hebben dropdown met alle beschikbare opties
- **Snelfilter knoppen**: Kleurgecodeerde knoppen voor directe filtering op:
  - Up-to-date (groen), Updates Wachtend (geel), Update Fouten (rood)
  - Verouderde OS (oranje), Handmatige Controle (paars), Non-Compliant (unieke rode kleur)
- **Filter Synchronisatie**: Automatische reset van conflicterende filters voor consistente ervaring
- **Dark Theme Support**: Optimale zichtbaarheid in zowel light als dark browser themes
- **Filter Counters**: Alle filter buttons tonen aantal machines per status

### 📊 **Enhanced Dashboard & Statistics**

- **Globale Statistieken**: Overzicht van alle tenants met totalen en percentages
- **Per-Client Statistieken**: Gedetailleerde breakdown per klant met visual cards
- **Compliance Percentages**: Up-to-date percentages en compliance ratios
- **Interactive Charts**: Grafische weergave van update status distributie
- **Professional Styling**: Bootstrap-compatible styling voor professionele uitstraling
- **Export Functies**: Volledige tabel export en gefilterde export opties

## 📋 Vereiste Microsoft Graph API Permissions

Voor volledige functionaliteit zijn de volgende **Application Permissions** vereist:

### 🔒 Device Management & Compliance

```text
DeviceManagementManagedDevices.Read.All
DeviceManagementConfiguration.Read.All
```

**Voor**: Device management API, compliance policy states, en device configuration informatie

### 🛡️ Security & Threat Hunting

```text
ThreatHunting.Read.All
```

**Voor**: Fallback device informatie via Advanced Hunting KQL queries

### 📊 Directory Information

```text
Device.Read.All
Directory.Read.All
```

**Voor**: Device directory informatie en organizational context

### ⚙️ Application Monitoring

```text
Application.Read.All
```

**Voor**: App Registration expiry monitoring en certificate status

## 🛠️ Installatie & Setup

### 1. Repository Setup

```powershell
git clone https://github.com/scns/Windows-Update-Report-MultiTenant.git
cd Windows-Update-Report-MultiTenant
```

### 2. Configuratie

```powershell
# Kopieer template bestanden
Copy-Item "_config.json" "config.json"
Copy-Item "_credentials.json" "credentials.json"

# Pas configuratie aan (zie CONFIG-UITLEG.md voor details)
notepad config.json
notepad credentials.json
```

### 3. App Registration Setup

1. Ga naar [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Maak nieuwe App Registration aan
3. Voeg de vereiste API permissions toe (zie lijst hierboven)
4. Genereer client secret
5. Vul `credentials.json` in met tenant ID, client ID en client secret per tenant

### 4. Eerste Run

```powershell
.\get-windows-update-report.ps1
```

Het script installeert automatisch benodigde PowerShell modules en genereert het eerste rapport.

## ⚙️ Configuratie Opties

### Timezone Configuration

```json
{
    "timezoneOffsetHours": 2
}
```

Opties:

- **Nederland (zomer)**: `2` (UTC+2)
- **Nederland (winter)**: `1` (UTC+1)
- **UTC tijd**: `0`
- **US Eastern**: `-5` (UTC-5)

### KB Mapping Cache

```json
{
    "kbMapping": {
        "cacheValidMinutes": 30,
        "timeoutSeconds": 10,
        "fallbackToLocalMapping": true
    }
}
```

### Export Management

```json
{
    "exportRetentionCount": 40,
    "cleanupOldExports": true,
    "autoOpenHtmlReport": false
}
```

## 📁 Bestandsstructuur

```text
Windows-Update-Report-MultiTenant/
├── get-windows-update-report.ps1    # Hoofd PowerShell script
├── config.json                      # Configuratie instellingen
├── credentials.json                  # Tenant credentials (exclusief git)
├── kb-mapping.json                   # Lokale KB mapping database
├── exports/                          # Gegenereerde rapporten
├── archive/                          # Gearchiveerde oude exports
├── backup/                           # Automatische backups
├── images/                           # Dashboard screenshots
├── CONFIG-UITLEG.md                  # Gedetailleerde configuratie uitleg
├── KB-CACHING-INFO.md               # KB caching documentatie
├── SECURITY.md                       # Beveiligingsbeleid
├── CONTRIBUTING.md                   # Contributie richtlijnen
└── CODE_OF_CONDUCT.md               # Gedragscode
```

## 🔧 Troubleshooting

### Permissions Errors

- Controleer of alle vereiste API permissions zijn toegekend
- Zorg ervoor dat permissions zijn "granted" door een admin
- Controleer client secret geldigheid

### Timezone Issues

- Pas `timezoneOffsetHours` aan in config.json
- Check of LastSeen tijden correct worden weergegeven
- Gebruik UTC offset voor uw tijdzone

### Cache Problems

- KB mapping cache wordt automatisch ververst na 30 minuten
- Bij problemen: verwijder `Global:CachedKBMapping` variabele
- Check internet connectiviteit voor online KB database

### Compliance Data Missing

- Zorg ervoor dat `DeviceManagementConfiguration.Read.All` permission is toegekend
- Controleer of devices enrolled zijn in Intune
- Fallback naar "Geen data" status is normaal voor niet-managed devices

## 📚 Documentatie Links

- **[Configuratie Uitleg](CONFIG-UITLEG.md)** - Gedetailleerde uitleg van alle config opties
- **[KB Caching Info](KB-CACHING-INFO.md)** - KB mapping cache configuratie en troubleshooting
- **[Security Policy](SECURITY.md)** - Beveiligingsbeleid en kwetsbaarheid rapportage
- **[Contributing Guidelines](CONTRIBUTING.md)** - Richtlijnen voor bijdragen aan het project
- **[Code of Conduct](CODE_OF_CONDUCT.md)** - Gedragscode voor contributors

## 🤝 Contributing

Bijdragen zijn welkom! Zie [CONTRIBUTING.md](CONTRIBUTING.md) voor richtlijnen.

## 📄 License

Dit project valt onder de [MIT License](LICENSE).

## 🔒 Security

Voor beveiligingsgerelateerde zaken, zie [SECURITY.md](SECURITY.md).

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/scns/Windows-Update-Report-MultiTenant/issues)
- **Email**: [info@maarten-schmeitz.nl](mailto:info@maarten-schmeitz.nl)
- **Documentation**: Zie de bijgevoegde MD bestanden voor gedetailleerde informatie

---

**Versie**: 3.0 | **Laatste Update**: Augustus 2025 | **PowerShell**: 7.2+ | **Microsoft Graph**: v1.0 & Beta

## Benodigdheden

- PowerShell 7.2+
- Microsoft Graph PowerShell SDK (wordt automatisch geïnstalleerd)
- Een Azure AD App Registration per tenant met de juiste permissies

## Voorbereiding

### 1. Maak een Azure AD App Registration aan

1. Ga naar [Azure Portal - App registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps).
2. Klik op **New registration** en geef de app een naam.
3. Na het aanmaken, ga naar **API permissions**.
4. **Verwijder alle standaard toegevoegde permissies** (zoals `User.Read`).
5. Voeg de volgende Microsoft Graph **Application** permissies toe:

#### Voor optimale functionaliteit (aanbevolen)

- `DeviceManagementManagedDevices.Read.All` - Voor gedetailleerde Windows Update informatie
- `ThreatHunting.Read.All` - Voor fallback functionaliteit
- `Application.Read.All` - Voor App Registration geldigheid monitoring

#### Minimale vereisten (fallback functionaliteit)

- `ThreatHunting.Read.All` - Voor basis Windows Update informatie
- `Application.Read.All` - Voor App Registration geldigheid monitoring

1. Klik op **Grant admin consent** voor deze permissies.
2. Ga naar **Certificates & secrets** en maak een nieuwe client secret aan. Noteer deze waarde direct.

> **💡 Tip**: Met `DeviceManagementManagedDevices.Read.All` krijg je specifieke KB nummers en gedetailleerde update informatie. Zonder deze permissie valt het script terug op basis functionaliteit via de Threat Hunting API.

### 2. Configureer het project

#### Credentials bestand

**Voor nieuwe installaties:**

1. Hernoem `_credentials.json` naar `credentials.json`
2. Vul de juiste waarden in voor je tenants

**Voor bestaande installaties:**

- Je bestaande `credentials.json` blijft werken zoals het is
- Geen wijzigingen nodig

Het `credentials.json` bestand heeft het volgende format:

```json
{
  "LoginCredentials": [
    {
      "ClientID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "Secret": "YOUR-CLIENT-SECRET",
      "TenantID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "customername": "KlantNaam",
      "color": "#1f77b4"
    }
    // Voeg meer tenants toe indien nodig
  ]
}
```

**color**: Geef per klant een vaste HTML kleurcode op (hex, bijvoorbeeld `#1f77b4`). Deze kleur wordt gebruikt in de grafieken van het HTML-dashboard.

#### Configuratie bestand

**Voor nieuwe installaties:**

1. Hernoem `_config.json` naar `config.json`
2. Pas de instellingen aan naar jouw behoeften

**Voor bestaande installaties:**

- Als je al een `config.json` hebt, voeg eventueel ontbrekende opties toe
- Vergelijk met `_config.json` om te zien welke nieuwe opties beschikbaar zijn

Het `config.json` bestand bevat alle instellingen:

```json
{
  "exportRetentionCount": 40,
  "cleanupOldExports": true,
  "exportDirectory": "exports",
  "archiveDirectory": "archive",
  "autoOpenHtmlReport": true,
  "lastSeenDaysFilter": 0,
  "kbMapping": {
    "kbMappingUrl": "https://mrtn.blog/wp-content/uploads/2025/08/kb-mapping.json",
    "timeoutSeconds": 10,
    "cacheValidMinutes": 30,
    "estimationThreshold": 1000,
    "showEstimationLabels": true,
    "fallbackToLocalMapping": true,
    "estimationLabels": {
      "buildDifference": "(geschat voor build {targetBuild})",
      "noMapping": "(geschat)",
      "oldMapping": "(verouderd)"
    }
  },
  "theme": {
    "default": "dark"
  }
}
```

**Configuratie opties:**

- `exportRetentionCount`: Aantal export bestanden dat behouden blijft per klant/type (oudere worden gearchiveerd)
- `cleanupOldExports`: Schakel automatische archivering in/uit (true/false)
- `exportDirectory`: Directory waar nieuwe export bestanden worden opgeslagen
- `archiveDirectory`: Directory waar oude export bestanden worden gearchiveerd
- `autoOpenHtmlReport`: Automatisch openen van HTML-rapport in webbrowser (true/false)
- `lastSeenDaysFilter`: Filtert de rapportage op basis van het aantal dagen sinds een device voor het laatst gezien is
- `kbMapping.kbMappingUrl`: URL naar online KB mapping database
- `kbMapping.timeoutSeconds`: Timeout voor online KB database requests (standaard: 10)
- `kbMapping.cacheValidMinutes`: Cache geldigheid in minuten (standaard: 30)
- `kbMapping.estimationThreshold`: Build verschil drempel voor KB estimaties
- `kbMapping.showEstimationLabels`: Toon labels voor geschatte KB nummers
- `kbMapping.fallbackToLocalMapping`: Gebruik lokale mapping als fallback
- `theme.default`: Standaard thema voor het dashboard ("dark" of "light")

> 📖 **Gedetailleerde configuratie uitleg**: Voor uitgebreide informatie over elke configuratie optie, zie [CONFIG-UITLEG.md](CONFIG-UITLEG.md)

### 3. Installeer benodigde PowerShell-modules

De benodigde modules worden automatisch geïnstalleerd bij het eerste gebruik van het script. Handmatige installatie is niet meer nodig.

## Gebruik

1. **Voor nieuwe installaties**: Hernoem `_credentials.json` naar `credentials.json` en `_config.json` naar `config.json`
2. **Voor bestaande installaties**: Controleer of je `config.json` alle benodigde opties bevat (vergelijk met `_config.json`)
3. Vul je tenant gegevens in het `credentials.json` bestand
4. Pas de instellingen in `config.json` aan naar jouw behoeften
5. Start het script:

```powershell
.\get-windows-update-report.ps1
```

1. Het script zal:
   - Automatisch benodigde modules installeren (indien nodig)
   - App Registration geldigheid controleren per tenant
   - Windows Update status ophalen via Device Management API (of fallback naar Threat Hunting API)
   - OS versie analyse uitvoeren en verouderde builds detecteren
   - Ontbrekende KB nummers identificeren voor machines met verouderde OS
   - CSV-bestanden genereren per klant met gedetailleerde informatie
   - Oude bestanden archiveren (indien geconfigureerd)
   - Een HTML-dashboard genereren met filterbare tabellen en snelfilters
   - Het rapport automatisch openen in je standaard webbrowser

2. De resultaten vind je in de geconfigureerde export directory, inclusief het interactieve HTML-dashboard.

## HTML Dashboard Functionaliteiten

### 📊 **Overzichtstabellen per Klant**

- **Filterbare DataTables**: Zoeken en sorteren op alle kolommen
- **Update Status dropdown**: Directe filtering op specifieke statussen
- **Snelfilter knoppen**: Kleurgecodeerde knoppen voor veelgebruikte filters
- **Export functionaliteit**: Exporteer volledige of gefilterde resultaten naar CSV

### 🎯 **Kolom Informatie**

| Kolom | Beschrijving |
|-------|-------------|
| **Device** | Computer naam |
| **Update Status** | Overall status (Up-to-date, Verouderde OS versie, etc.) |
| **Missing Updates** | Specifieke KB nummers en update namen die ontbreken |
| **Details** | Statusberichten en diagnostische informatie |
| **OS Version** | Windows build versie |
| **KB Method** | Bron van KB informatie (Online, Cache, ExpiredCache, Local, Estimated) |
| **Count** | Binary indicator (0 = OK, 1 = aandacht vereist) |
| **LastSeen** | Laatste synchronisatie datum |
| **LoggedOnUsers** | Huidige gebruikers |

### 🔍 **Filter Functionaliteiten**

```html
<!-- Snelfilter knoppen -->
🔘 Alle statussen    🟢 Up-to-date    🟠 Verouderde OS    
🔴 Handmatige controle    🔵 Waarschijnlijk up-to-date    🟣 Errors
```

### 📈 **Grafieken en Analyses**

- **Count trend per dag**: Laat zien hoe de Windows Update situatie evolueert
- **Per klant overzicht**: Vergelijk update status tussen verschillende tenants
- **OS versie analyse**: Percentage verdeling van Windows builds

### 🌙 **Thema Ondersteuning**

- **Dark/Light mode toggle**: Schakel eenvoudig tussen donker en licht thema
- **Configureerbare standaard**: Stel je voorkeur in via `config.json`
- **Gebruiksvriendelijk**: FontAwesome iconen voor consistente weergave

### 🗄️ **KB Mapping Database Dashboard**

Het HTML-dashboard bevat een dedicated "KB Mapping Database" tabblad dat uitgebreide informatie toont over de KB mapping database:

#### Database Status Overzicht

- **Database Status**: ✅ Beschikbaar / ❌ Niet beschikbaar
- **Bron Methode**: Online, Cache, ExpiredCache, Error, of Exception
- **Totaal Entries**: Aantal beschikbare KB mappings in de database
- **Laatste Update**: Timestamp van laatste cache refresh

#### KB Mapping Tabel

**Volledige database weergave** met filterbare/sorteerbare kolommen:

- **Build Number**: Windows OS build nummer (bijv. 26100, 22631)
- **KB Number**: Corresponderende KB update (bijv. KB5041585)
- **Update Title**: Beschrijving van de update
- **Release Date**: Officiële release datum van de update
- **OS Version**: Windows versie categorie (Windows 11 24H2, Windows 10 22H2, Historical)

#### Cache Intelligence Features

- **Real-time status**: Toont huidige cache status en bron van informatie
- **Performance metrics**: Zichtbaarheid in cache effectiviteit
- **Fallback transparency**: Duidelijke indicatie wanneer fallback wordt gebruikt
- **Historical data**: Toegang tot historische KB mappings per jaar

## API Methodologie en Fallback

### 🎯 **Primary Method: Device Management API**

**Vereist**: `DeviceManagementManagedDevices.Read.All` permissie

**Voordelen**:

- Specifieke KB nummers van ontbrekende updates
- Detailed compliance informatie  
- Windows Update state tracking
- Configuration policy violations

**Voorbeeld output**:

```text
Missing Updates: "2024-08 Cumulative Update voor Windows (KB5041585 of nieuwer)"
Details: "Windows Update status: Recent gesynchroniseerd, geen problemen"
```

### 🔄 **Fallback Method: Threat Hunting API**

**Vereist**: `ThreatHunting.Read.All` permissie

**Gebruikt wanneer**:

- Device Management API niet beschikbaar
- Onvoldoende permissies voor Device Management
- Tenant heeft geen Intune licenties

**Output**:

```text
Missing Updates: (leeg - geen specifieke KB info beschikbaar)
Details: "Windows Update status: Controleer handmatig - Device niet in Intune beheer"
```

### 🧠 **Intelligente OS Analyse**

**Voor alle scenarios**:

Ongeacht welke API gebruikt wordt, het script analyseert OS versies en:

- Detecteert nieuwste Windows build in de omgeving
- Identificeert machines met verouderde builds  
- Suggereert specifieke KB updates voor bekende build verschillen
- Geeft praktische aanbevelingen voor IT beheerders

**Voorbeeld voor verouderde OS**:

```text
Missing Updates: "Waarschijnlijk ontbrekende cumulative update (build verschil: 294); 2024-08 Cumulative Update voor Windows (KB5041585 of nieuwer)"
Update Status: "Verouderde OS versie"
```

## Bestandsstructuur

### Template bestanden (meegeleverd)

```text
Windows-Update-Report-MultiTenant/
├── _credentials.json     # Template voor credentials (hernoem naar credentials.json)
├── _config.json         # Template voor configuratie (hernoem naar config.json)
├── CONFIG-UITLEG.md     # Gedetailleerde configuratie uitleg
├── get-windows-update-report.ps1
└── readme.md
```

### Na configuratie en uitvoering

```text
Windows-Update-Report-MultiTenant/
├── _credentials.json     # Template bestand (blijft bestaan)
├── _config.json         # Template bestand (blijft bestaan)
├── credentials.json     # Jouw tenant configuratie
├── config.json          # Jouw instellingen
├── CONFIG-UITLEG.md
├── get-windows-update-report.ps1
├── exports/             # Configureerbare export directory
│   ├── 20250822_KlantA_Windows_Update_report_Overview.csv
│   ├── 20250822_KlantA_Windows_Update_report_ByUpdate.csv
│   ├── 20250822_KlantB_Windows_Update_report_Overview.csv
│   ├── 20250822_KlantB_Windows_Update_report_ByUpdate.csv
│   └── Windows_Update_Overview.html
└── archive/             # Oude bestanden worden hier gearchiveerd
    ├── 20250821_KlantA_Windows_Update_report_Overview.csv
    └── ... (oudere bestanden)
```

## Nieuwe functies in v3.0

### 🎯 **KB Detection & Update Intelligence**

- **Specifieke KB nummers**: Identificeert ontbrekende updates zoals KB5041585
- **Build gap analyse**: Analyseert verschil tussen huidige en nieuwste OS builds
- **Smart suggestions**: Geeft praktische aanbevelingen op basis van build verschillen
- **Multiple API support**: Device Management API met Threat Hunting fallback

### �️ **KnowledgeBase Mapping Database & Intelligent Caching**

- **Online KB database**: Uitgebreide externe database met Windows Update KB mappings
- **Intelligent caching**: Downloads database eenmalig per sessie, cached voor 30 minuten (configureerbaar)
- **Performance optimalisatie**: Vermindert server load met 95%+ door slim cache beheer
- **Fallback mechanisme**: Gebruikt expired cache bij netwerk problemen voor maximale betrouwbaarheid
- **Cache transparantie**: KB Method kolom toont bron van elke KB lookup (Online/Cache/ExpiredCache/Local/Estimated)
- **Multi-platform database**: Ondersteunt Windows 10, Windows 11, en historische versies met 27+ KB mappings

- **Update Status dropdown**: Vervang tekstfilter met dropdown voor exacte filtering
- **Snelfilter knoppen**: Kleurgecodeerde knoppen voor directe access tot veelgebruikte filters
- **Synchronized filtering**: Dropdown en snelfilters werken samen voor optimale UX
- **Filter persistence**: Behoud filter instellingen tijdens sessie

### 📊 **Enhanced Data Presentation**

- **Separated columns**: "Missing Updates" voor KB nummers, "Details" voor statusberichten
- **Improved Count logic**: Binary indicator (0/1) voor duidelijke status indicatie
- **OS version analysis**: Percentage breakdown van Windows builds in omgeving
- **Status categorization**: Duidelijke update status categorieën voor betere insights

## Backup functionaliteit

Het script ondersteunt automatische back-ups van exports, archief en configuratiebestanden:

- **Export back-up**: Maakt een zip-bestand van de exports directory
- **Archief back-up**: Maakt een zip-bestand van de archive directory
- **Config back-up**: Maakt een zip-bestand van config.json en credentials.json
- **Retentie**: Het aantal te bewaren back-ups is instelbaar per type
- **Configuratie**: Alle paden en instellingen zijn te beheren via `config.json`

### Configuratie opties

In het `config.json` bestand kun je per back-up type instellen of deze actief is en hoeveel back-ups bewaard blijven:

```json
"backup": {
    "enableExportBackup": true,
    "enableArchiveBackup": true,
    "enableConfigBackup": true,
    "exportBackupRetention": 5,
    "archiveBackupRetention": 5,
    "configBackupRetention": 5,
    "backupRoot": "backup",
    "exportBackupSubfolder": "export_backup",
    "archiveBackupSubfolder": "archive_backup",
    "configBackupSubfolder": "config_backup"
}
```

## App Registration Status Dashboard

Het HTML-dashboard bevat een speciaal "App Registrations" tabblad dat een overzicht toont van alle client secret vervaldatums:

### App Registration Dashboard Functionaliteit

- **Centraal overzicht**: Alle tenants en hun App Registration status in één tabel
- **Status indicatoren**: Kleurgecodeerde waarschuwingen voor vervaldatums:
  - 🟢 **Groen**: Meer dan 30 dagen geldig
  - 🟠 **Oranje**: 7-30 dagen tot verval (waarschuwing)
  - 🔴 **Rood**: Minder dan 7 dagen tot verval (kritiek) of geen toegang
- **Gedetailleerde informatie**:
  - Customer naam
  - Status bericht
  - Dagen tot verval
  - Exacte vervaldatum (DD-MM-YYYY)
- **Filterbare tabel**: Zoeken en sorteren op alle kolommen via DataTables
- **Export functionaliteit**: Mogelijkheid tot CSV export van de status gegevens

## Troubleshooting

### ❌ **"Device Management API niet beschikbaar"**

**Oorzaak**: Ontbrekende `DeviceManagementManagedDevices.Read.All` permissie of geen Intune licenties  
**Oplossing**: Script valt automatisch terug op Threat Hunting API - geen actie vereist

### ❌ **"Geen geldige client secret gevonden"**

**Oorzaak**: Ontbrekende `Application.Read.All` permissie  
**Oplossing**: Voeg permissie toe in Azure Portal en verleen admin consent

### ❌ **Missing Updates kolom is leeg**

**Oorzaak**: Machines zijn werkelijk up-to-date of gebruiken fallback API  
**Verwachting**: Dit is normaal - machines met verouderde OS krijgen automatisch KB suggesties

### ✅ **Permissie Verificatie**

Controleer in Azure Portal of je App Registration deze permissies heeft:

- ✅ `DeviceManagementManagedDevices.Read.All` (voor KB nummers)
- ✅ `ThreatHunting.Read.All` (voor fallback)  
- ✅ `Application.Read.All` (voor App Registration monitoring)

## Opmerkingen

- Het script werkt optimaal met Device Management permissies, maar functioneert ook met alleen Threat Hunting permissies
- KB nummers worden alleen getoond bij Device Management API toegang of bij gedetecteerde verouderde OS versies
- Voor meer informatie over App Registrations, zie de [Microsoft Docs](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- Het script ondersteunt zowel Windows 10 als Windows 11 omgevingen

---

© 2025 l 22/08/2025 by Maarten Schmeitz

[commits-shield]: https://img.shields.io/github/commit-activity/m/scns/Windows-Update-Report-MultiTenant.svg
[commits]: https://github.com/scns/Windows-Update-Report-MultiTenant/commits/main
[github-last-commit]: https://img.shields.io/github/last-commit/scns/Windows-Update-Report-MultiTenant.svg?style=plasticr
[github-master]: https://github.com/scns/Windows-Update-Report-MultiTenant/commits/main
[license-shield]: https://img.shields.io/github/license/scns/Windows-Update-Report-MultiTenant.svg
[contributors-url]: https://github.com/scns/Windows-Update-Report-MultiTenant/graphs/contributors
[contributors-shield]: https://img.shields.io/github/contributors/scns/Windows-Update-Report-MultiTenant.svg
[forks-shield]: https://img.shields.io/github/forks/scns/Windows-Update-Report-MultiTenant.svg
[forks-url]: https://github.com/scns/Windows-Update-Report-MultiTenant/network/members
[stars-shield]: https://img.shields.io/github/stars/scns/Windows-Update-Report-MultiTenant.svg
[stars-url]: https://github.com/scns/Windows-Update-Report-MultiTenant/stargazers
[issues-shield]: https://img.shields.io/github/issues/scns/Windows-Update-Report-MultiTenant.svg
[issues-url]: https://github.com/scns/Windows-Update-Report-MultiTenant/issues
