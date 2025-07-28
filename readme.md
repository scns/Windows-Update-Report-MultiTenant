# Windows Update Report MultiTenant

Dit PowerShell-project genereert een overzichtsrapport van ontbrekende Windows-updates per device voor meerdere tenants via Microsoft Graph. Het resultaat is een dynamisch HTML-dashboard met filterbare tabellen en grafieken.

## Functionaliteit

- Haalt per tenant de ontbrekende Windows-updates op via Microsoft Graph Threat Hunting API.
- Exporteert resultaten naar CSV-bestanden per klant.
- Genereert een HTML-dashboard met filterbare tabellen (DataTables) en grafieken (Chart.js).
- Ondersteunt meerdere tenants via een `credentials.json`-bestand.

## Benodigdheden

- PowerShell 5+
- Microsoft Graph PowerShell SDK (`Install-Module Microsoft.Graph`)
- Een Azure AD App Registration per tenant met de juiste permissies

## Voorbereiding

### 1. Maak een Azure AD App Registration aan

1. Ga naar [Azure Portal - App registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps).
2. Klik op **New registration** en geef de app een naam.
3. Na het aanmaken, ga naar **API permissions**.
4. **Verwijder alle standaard toegevoegde permissies** (zoals `User.Read`).
5. Voeg de volgende Microsoft Graph **Application** permissies toe:
    - `SecurityEvents.Read.All`
    - `ThreatHunting.Read.All`
6. Klik op **Grant admin consent** voor deze permissies.
7. Ga naar **Certificates & secrets** en maak een nieuwe client secret aan. Noteer deze waarde direct.

### 2. Vul het `credentials.json`-bestand

Maak een bestand `credentials.json` aan in de root van dit project met het volgende format:

```json
{
  "LoginCredentials": [
    {
      "ClientID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "Secret": "YOUR-CLIENT-SECRET",
      "TenantID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "customername": "KlantNaam"
    }
    // Voeg meer tenants toe indien nodig
  ]
}
```

### 3. Installeer benodigde PowerShell-modules

Open PowerShell als administrator en voer uit:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Gebruik

1. Plaats je `credentials.json` in de projectmap.
2. Start het script:

```powershell
.\get-windows-update-report.ps1
```

3. De resultaten vind je in de map `exports`, inclusief een HTML-dashboard (`Windows_Update_Overview.html`).

## Opmerkingen

- Zorg dat je app registration alleen de genoemde permissies bevat.
- Het script werkt alleen met tenants waar de app registration en rechten correct zijn ingesteld.
- Voor meer informatie over App Registrations, zie de [Microsoft Docs](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).

---

© 2025 by Maarten Schmeitz
