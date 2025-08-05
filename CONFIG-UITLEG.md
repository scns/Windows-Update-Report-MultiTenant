# Configuratie Uitleg - config.json

Dit bestand legt uit wat elke instelling in `config.json` doet:

## Configuratie Parameters

### `exportRetentionCount`

- **Type**: Number
- **Standaard**: 40
- **Beschrijving**: Het aantal export bestanden dat behouden blijft per klant en per type (Overview/ByUpdate). Bijvoorbeeld: als deze waarde op 10 staat, worden de laatste 10 "Overview" bestanden van elke klant behouden, en de laatste 10 "ByUpdate" bestanden van elke klant. Oudere bestanden worden automatisch naar de archief directory verplaatst.

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

## Voorbeeld Configuratie

```json
{
    "exportRetentionCount": 40,
    "cleanupOldExports": true,
    "exportDirectory": "exports",
    "archiveDirectory": "archive"
}
```

Met deze instellingen:

- De laatste 40 bestanden per klant/type blijven in de "exports" folder
- Oudere bestanden worden automatisch verplaatst naar de "archive" folder
- Nieuwe exports gaan naar de "exports" directory
- Gearchiveerde bestanden gaan naar de "archive" directory
