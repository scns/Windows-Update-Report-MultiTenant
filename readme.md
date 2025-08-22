# Windows Update Report MultiTenant

| Repository Status | Windows Update Report |
| :--- | :--- |
|  [![last commit time][github-last-commit]][github-master] [![GitHub Activity][commits-shield]][commits] | |
| [![License][license-shield]](LICENSE) [![Forks][forks-shield]][forks-url] [![Stargazers][stars-shield]][stars-url] [![Issues][issues-shield]][issues-url] | [![Contributors][contributors-shield]][contributors-url] [![GitHub release](https://img.shields.io/github/release/scns/Windows-Update-Report-MultiTenant.svg)](https://GitHub.com/scns/Windows-Update-Report-MultiTenant/releases) |

![Dashboard voorbeeld](images/001.png)

Dit PowerShell-project genereert een overzichtsrapport van ontbrekende Windows-updates per device voor meerdere tenants via Microsoft Graph. Het resultaat is een dynamisch HTML-dashboard met filterbare tabellen en grafieken.

## Functionaliteit

- **Automatische module installatie**: Controleert en installeert automatisch benodigde PowerShell modules
- **Configureerbare instellingen**: Alle instellingen beheerbaar via `config.json`
- **Multi-tenant ondersteuning**: Haalt per tenant de ontbrekende Windows-updates op via Microsoft Graph Threat Hunting API
- **Flexibele export opties**: Exporteert resultaten naar CSV-bestanden per klant
- **Interactief HTML-dashboard**: Genereert een dashboard met filterbare tabellen (DataTables) en grafieken (Chart.js)
- **App Registration status monitoring**: Dedicated tabblad voor overzicht van client secret vervaldatums per tenant
- **Intelligente bestandsbeheer**: Automatische archivering van oude export bestanden
- **Automatische browser integratie**: Configureerbaar automatisch openen van het gegenereerde rapport in de standaard webbrowser
- **App Registration monitoring**: Controleert automatisch de geldigheid van client secrets en waarschuwt voor vervaldatums
- **Automatische back-up**: Maakt automatisch back-ups van exports, archief en configuratiebestanden

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
    - `SecurityEvents.Read.All` - Voor het lezen van beveiligingsgegevens
    - `ThreatHunting.Read.All` - Voor het uitvoeren van threat hunting queries
    - `Application.Read.All` - Voor het controleren van App Registration geldigheid (client secret vervaldatums)
6. Klik op **Grant admin consent** voor deze permissies.
7. Ga naar **Certificates & secrets** en maak een nieuwe client secret aan. Noteer deze waarde direct.

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
  "lastSeenDaysFilter": 0
}
```

**Configuratie opties:**

- `exportRetentionCount`: Aantal export bestanden dat behouden blijft per klant/type (oudere worden gearchiveerd)
- `cleanupOldExports`: Schakel automatische archivering in/uit (true/false)
- `exportDirectory`: Directory waar nieuwe export bestanden worden opgeslagen
- `archiveDirectory`: Directory waar oude export bestanden worden gearchiveerd
- `autoOpenHtmlReport`: Automatisch openen van HTML-rapport in webbrowser (true/false)
- `lastseenDaysFilter` : Filtert de rapportage (tabel én grafiek) op basis van het aantal dagen sinds een device voor het laatst gezien is.

> 📖 **Gedetailleerde configuratie uitleg**: Voor uitgebreide informatie over elke configuratie optie, zie [CONFIG-UITLEG.md](CONFIG-UITLEG.md)

### 3. Installeer benodigde PowerShell-modules

De benodigde modules worden automatisch geïnstalleerd bij het eerste gebruik van het script. Handmatige installatie is niet meer nodig.

## Gebruik

1. **Voor nieuwe installaties**: Hernoem `_credentials.json` naar `credentials.json` en `_config.json` naar `config.json`
1. **Voor bestaande installaties**: Controleer of je `config.json` alle benodigde opties bevat (vergelijk met `_config.json`)
1. Vul je tenant gegevens in het `credentials.json` bestand
1. Pas de instellingen in `config.json` aan naar jouw behoeften
1. Start het script:

```powershell
.\get-windows-update-report.ps1
```

1. Het script zal:
   - Automatisch benodigde modules installeren (indien nodig)
   - App Registration geldigheid controleren per tenant (client secret vervaldatums)
   - Data ophalen van alle geconfigureerde tenants
   - CSV-bestanden genereren per klant (inclusief App Registration status)
   - Oude bestanden archiveren (indien geconfigureerd)
   - Een HTML-dashboard genereren
   - Het rapport automatisch openen in je standaard webbrowser

1. De resultaten vind je in de geconfigureerde export directory, inclusief het HTML-dashboard.

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
│   ├── 20250806_KlantA_Windows_Update_report_Overview.csv
│   ├── 20250806_KlantA_Windows_Update_report_ByUpdate.csv
│   ├── 20250806_KlantB_Windows_Update_report_Overview.csv
│   ├── 20250806_KlantB_Windows_Update_report_ByUpdate.csv
│   └── Windows_Update_Overview.html
└── archive/             # Oude bestanden worden hier gearchiveerd
    ├── 20250805_KlantA_Windows_Update_report_Overview.csv
    └── ... (oudere bestanden)
```

## Nieuwe functies in v2.0

- **Automatische module installatie**: Geen handmatige module installatie meer nodig
- **Configureerbare archivering**: Oude bestanden worden verplaatst naar archief in plaats van verwijderd
- **Flexibele directory instellingen**: Configureerbare export en archief directories
- **Automatische browser integratie**: HTML rapport wordt automatisch geopend
- **App Registration monitoring**: Automatische controle van client secret vervaldatums met kleurgecodeerde waarschuwingen
- **Verbeterde feedback**: Kleurgecodeerde status berichten tijdens uitvoering
- **Intelligente bestandsbeheer**: Configureerbaar aantal bestanden dat behouden blijft

## Backup functionaliteit

Het script ondersteunt automatische back-ups van exports, archief en configuratiebestanden:

- **Export back-up**: Maakt een zip-bestand van de exports directory
- **Archief back-up**: Maakt een zip-bestand van de archive directory
- **Config back-up**: Maakt een zip-bestand van config.json en credentials.json
- **Retentie**: Het aantal te bewaren back-ups is instelbaar per type
- **Configuratie**: Alle paden en instellingen zijn te beheren via `config.json`

### Configuratie opties

In het `config.json` bestand kun je per back-up type instellen of deze actief is en hoeveel back-ups bewaard blijven. Je kunt ook de root en subfolders voor back-ups aanpassen:

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

Back-ups worden opgeslagen in de opgegeven subfolders onder de root. Overtollige back-ups worden automatisch verwijderd.

## Thema (dark/light mode)

Het HTML-dashboard ondersteunt een dark/light modus. Je kunt de standaardmodus instellen via de configuratie:

```json
"theme": {
    "default": "dark" // opties: "dark", "light"
}
```

- **default**: Kies "dark" voor standaard donkere modus, of "light" voor standaard lichte modus.
- Je kunt altijd handmatig wisselen via de knop rechtsboven in het dashboard.

## Dark mode toggle

Het HTML-dashboard bevat een dark mode toggle-knop rechtsboven. Hiermee kun je eenvoudig wisselen tussen licht en donker thema. De knop gebruikt FontAwesome iconen (zon/maan) voor een consistente weergave in alle browsers.

- **Werking:** Klik op de knop om te wisselen tussen light en dark mode.
- **Iconen:** De knop toont automatisch een zon (light mode) of maan (dark mode) icoon.
- **Compatibiliteit:** Door gebruik van FontAwesome worden de iconen overal correct weergegeven.

Deze functie verbetert de leesbaarheid en gebruikerservaring, vooral bij gebruik in donkere omgevingen.

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

### Toegangsvereisten

Voor volledige functionaliteit moet elke App Registration de `Application.Read.All` permissie hebben:

1. Ga naar [Azure Portal - App registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)
2. Selecteer je App Registration
3. Ga naar **API permissions**
4. Voeg `Application.Read.All` toe (Microsoft Graph Application permission)
5. Klik op **Grant admin consent**

**Zonder deze permissie** toont het dashboard "Geen geldige client secret gevonden" voor die tenant.

## Opmerkingen

- Zorg dat je app registration alleen de genoemde permissies bevat.
- Het script werkt alleen met tenants waar de app registration en rechten correct zijn ingesteld.
- Voor meer informatie over App Registrations, zie de [Microsoft Docs](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).

---

© 2025 l 15/08/2025 by Maarten Schmeitz

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
