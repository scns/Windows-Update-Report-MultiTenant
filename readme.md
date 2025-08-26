# Windows Update Report MultiTenant

| Repository Status | Windows Update Report |
| :--- | :--- |
|  [![last commit time][github-last-commit]][github-master] [![GitHub Activity][commits-shield]][commits] | |
| [![License][license-shield]](LICENSE) [![Forks][forks-shield]][forks-url] [![Stargazers][stars-shield]][stars-url] [![Issues][issues-shield]][issues-url] | [![Contributors][contributors-shield]][contributors-url] [![GitHub release](https://img.shields.io/github/release/scns/Windows-Update-Report-MultiTenant.svg)](https://GitHub.com/scns/Windows-Update-Report-MultiTenant/releases) |

![Dashboard voorbeeld](images/001.png)

Dit PowerShell-project genereert een uitgebreid overzichtsrapport van Windows Update status per device voor meerdere tenants via Microsoft Graph. Het resultaat is een dynamisch HTML-dashboard met filterbare tabellen, grafieken en gedetailleerde KB informatie.

## Hoofdfunctionaliteit

- **Automatische module installatie**: Controleert en installeert automatisch benodigde PowerShell modules
- **Configureerbare instellingen**: Alle instellingen beheerbaar via `config.json`
- **Multi-tenant ondersteuning**: Haalt per tenant Windows Update status op via Microsoft Graph Device Management API met Threat Hunting API fallback
- **Intelligente KB detectie**: Toont specifieke ontbrekende KB nummers en security patches
- **OS versie analyse**: Automatische detectie van verouderde Windows builds en aanbevelingen
- **Flexibele export opties**: Exporteert resultaten naar CSV-bestanden per klant
- **Interactief HTML-dashboard**: Genereert dashboard met filterbare tabellen, snelfilters en grafieken
- **App Registration status monitoring**: Dedicated tabblad voor overzicht van client secret vervaldatums per tenant
- **Intelligente bestandsbeheer**: Automatische archivering van oude export bestanden
- **Automatische browser integratie**: Configureerbaar automatisch openen van het gegenereerde rapport

## Nieuwe Functionaliteiten v3.0

### 🎯 **Intelligente Update Detectie**

- **Specifieke KB nummers**: Toont ontbrekende KB updates zoals "KB5041585" voor machines met verouderde OS
- **Build analyse**: Analyseert OS versie verschillen en suggereert benodigde cumulative updates
- **Update status categorieën**:
  - "Up-to-date", "Verouderde OS versie", "Handmatige controle vereist"
  - "Waarschijnlijk up-to-date", "Updates wachtend", "Update fouten"

### 🔍 **Geavanceerde Filtering**

- **Dropdown filters**: Update Status kolom heeft dropdown met alle beschikbare statussen
- **Snelfilter knoppen**: Kleurgecodeerde knoppen voor directe filtering op:
  - 🟢 Up-to-date
  - 🟠 Verouderde OS versie  
  - 🔴 Handmatige controle vereist
  - 🔵 Waarschijnlijk up-to-date
  - 🟣 Errors
- **Gesynchroniseerde filtering**: Dropdown en snelfilters werken samen

### 📊 **Uitgebreide Rapportage**

- **Missing Updates kolom**: Toont specifieke ontbrekende KB nummers en update namen
- **Details kolom**: Bevat statusberichten en diagnostische informatie  
- **Count logica**: Binary indicator (0 = up-to-date, 1 = aandacht vereist)
- **OS versie trending**: Toont percentage machines op nieuwste vs verouderde builds

### 🔄 **API Intelligentie**

- **Primary**: Device Management API voor gedetailleerde Windows Update informatie
- **Secondary**: Configuration compliance voor policy violations
- **Fallback**: Threat Hunting API voor tenants zonder Device Management toegang
- **Smart detection**: Build analysis voor praktische update suggesties

## Benodigdheden

- PowerShell 5+
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

### 🔍 **Geavanceerde Filtering Interface**

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
