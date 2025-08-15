# Configuratie Uitleg - config.json

Dit bestand legt uit wat elke instelling in `config.json` doet.

## Setup

**Voor nieuwe installaties:**

- Hernoem `_config.json` naar `config.json`
- Pas de waarden aan naar jouw behoeften

**Voor bestaande installaties:**

- Controleer of je bestaande `config.json` alle onderstaande opties bevat
- Voeg ontbrekende opties toe indien nodig
- Vergelijk met `_config.json` voor de nieuwste template

## Configuratie Parameters

### `exportRetentionCount`

- **Type**: Number
- **Standaard**: 40
- **Beschrijving**: Het aantal export bestanden dat behouden blijft per klant en per type (Overview/ByUpdate). Bijvoorbeeld: als deze waarde op 40 staat, worden de laatste 40 "Overview" bestanden van elke klant behouden, en de laatste 40 "ByUpdate" bestanden van elke klant. Oudere bestanden worden automatisch naar de archief directory verplaatst.

### `cleanupOldExports`

- **Type**: Boolean (true/false)
- **Standaard**: true
- **Beschrijving**: Schakelt de automatische archivering van oude export bestanden in of uit.
  - `true`: Oude bestanden worden verplaatst naar de archief directory
  - `false`: Alle bestanden blijven in de export directory (geen automatische archivering)

### `exportDirectory`

- **Type**: String
- **Standaard**: "exports"
- **Beschrijving**: De naam van de directory waar nieuwe export bestanden worden opgeslagen. Dit pad is relatief ten opzichte van de script locatie. De directory wordt automatisch aangemaakt als deze niet bestaat.

### `archiveDirectory`

- **Type**: String
- **Standaard**: "archive"
- **Beschrijving**: De naam van de directory waar oude export bestanden naartoe worden verplaatst wanneer `cleanupOldExports` is ingeschakeld. Dit pad is relatief ten opzichte van de script locatie. De directory wordt automatisch aangemaakt als deze niet bestaat.

### `autoOpenHtmlReport`

- **Type**: Boolean (true/false)
- **Standaard**: true
- **Beschrijving**: Bepaalt of het gegenereerde HTML-rapport automatisch wordt geopend in de standaard webbrowser nadat het script is voltooid.
  - `true`: Het HTML-rapport wordt automatisch geopend in de standaard webbrowser
  - `false`: Het HTML-rapport wordt niet automatisch geopend (handig voor server omgevingen of geautomatiseerde runs)

### `lastSeenDaysFilter`

- **Type**: Number
- **Standaard**: 0
- **Beschrijving**: Filtert de rapportage (tabel én grafiek) op basis van het aantal dagen sinds een device voor het laatst gezien is.
  - `0`: Alle devices worden weergegeven
  - `>0`: Alleen devices die in de laatste N dagen zijn gezien worden getoond

Deze filtering is zichtbaar in zowel de tabellen als de grafieken in het HTML-dashboard.

### `backup`

- **Type**: Object
- **Beschrijving**: Instellingen voor automatische back-ups van exports, archief en configuratiebestanden.

#### `enableExportBackup`

- **Type**: Boolean
- **Beschrijving**: Schakelt back-up van de exports directory in/uit.

#### `enableArchiveBackup`

- **Type**: Boolean
- **Beschrijving**: Schakelt back-up van de archive directory in/uit.

#### `enableConfigBackup`

- **Type**: Boolean
- **Beschrijving**: Schakelt back-up van config.json en credentials.json in/uit.

#### `exportBackupRetention`, `archiveBackupRetention`, `configBackupRetention`

- **Type**: Number
- **Beschrijving**: Aantal te bewaren back-ups per type. Overtollige back-ups worden automatisch verwijderd.

#### `backupRoot`, `exportBackupSubfolder`, `archiveBackupSubfolder`, `configBackupSubfolder`

- **Type**: String
- **Beschrijving**: Paden voor de root en subfolders van back-ups. Je kunt deze aanpassen naar wens.

### `theme`

- **Type**: Object
- **Beschrijving**: Instellingen voor het standaard dashboard-thema.

#### `default`

- **Type**: String
- **Opties**: "dark", "light"
- **Beschrijving**: Bepaalt of het dashboard standaard in donkere of lichte modus start. Je kunt altijd handmatig wisselen via de knop in het dashboard.

**Voorbeeld:**

```json
"theme": {
    "default": "dark"
}
```

## Voorbeeld Configuratie

```json
{
    "exportRetentionCount": 40,
    "cleanupOldExports": true,
    "exportDirectory": "exports",
    "archiveDirectory": "archive",
    "autoOpenHtmlReport": true,
    "lastSeenDaysFilter": 0,
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
    },
    "theme": {
        "default": "dark"
    }
}
```

Met deze instellingen:

- De laatste 40 bestanden per klant/type blijven in de "exports" folder
- Oudere bestanden worden automatisch verplaatst naar de "archive" folder
- Nieuwe exports gaan naar de "exports" directory
- Gearchiveerde bestanden gaan naar de "archive" directory
- Het HTML-rapport wordt automatisch geopend in de webbrowser
- De rapportage wordt gefilterd op basis van het aantal dagen sinds een device voor het laatst gezien is (indien ingesteld)
- Back-ups worden gemaakt van de exports, archief en configuratiebestanden volgens de opgegeven instellingen
- Het dashboard start standaard in donkere modus (kan handmatig gewijzigd worden)
