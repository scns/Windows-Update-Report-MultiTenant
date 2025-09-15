# Credentials Setup Guide v3.0

Deze handleiding legt uit hoe je de credentials correct instelt voor het Windows Update Report MultiTenant script.

## 📋 Overzicht

Het script gebruikt Azure App Registrations om verbinding te maken met Microsoft Graph API per tenant/klant. Elke klant heeft zijn eigen App Registration nodig met specifieke permissions.

## 🚀 Quick Start

### 1. Template Setup

```powershell
# Kopieer de template naar het werkbestand
Copy-Item "_credentials.json" "credentials.json"
```

### 2. Basis Structuur

Het `credentials.json` bestand heeft de volgende structuur:

```json
{
"LoginCredentials":[
    {"customername":"Customer1", "ClientID": "[Your Client/App ID]", "Secret":"[Your App Secret]", "TenantID": "[Your Tenant ID]", "color": "#1f77b4"},
    {"customername":"Customer2", "ClientID": "[Your Client/App ID]", "Secret":"[Your App Secret]", "TenantID": "[Your Tenant ID]", "color": "#ff7f0e"}
]
}
```

## 🔧 Azure App Registration Setup

Voor elke klant/tenant heb je een App Registration nodig. Volg deze stappen per klant:

### Stap 1: App Registration Aanmaken

1. Ga naar [Azure Portal](https://portal.azure.com)
2. Navigeer naar **Azure Active Directory** → **App registrations**
3. Klik op **"New registration"**
4. Vul in:
   - **Name**: `Windows Update Report - [KlantNaam]`
   - **Supported account types**: "Accounts in this organizational directory only"
   - **Redirect URI**: Laat leeg
5. Klik **"Register"**

### Stap 2: Client Secret Genereren

1. Ga naar je nieuwe App Registration
2. Klik op **"Certificates & secrets"**
3. Onder **"Client secrets"**, klik **"New client secret"**
4. Vul in:
   - **Description**: `Windows Update Report Secret`
   - **Expires**: 24 months (aanbevolen)
5. Klik **"Add"**
6. **⚠️ BELANGRIJK**: Kopieer de **Value** direct! Deze wordt maar één keer getoond.

### Stap 3: API Permissions Toevoegen

1. Ga naar **"API permissions"**
2. Klik **"Add a permission"** → **"Microsoft Graph"** → **"Application permissions"**
3. Voeg de volgende permissions toe:

#### 🔒 Device Management & Compliance

```text
DeviceManagementManagedDevices.Read.All
DeviceManagementConfiguration.Read.All
```

#### 🛡️ Security & Threat Hunting

```text
ThreatHunting.Read.All
```

#### 📊 Directory Information

```text
Device.Read.All
Directory.Read.All
```

#### ⚙️ Application Monitoring

```text
Application.Read.All
```

1. Klik **"Add permissions"**
2. **⚠️ KRITIEK**: Klik **"Grant admin consent for [Organization]"** en bevestig

### Stap 4: Gegevens Verzamelen

Verzamel de volgende gegevens van je App Registration:

1. **Tenant ID**: Azure Portal → Azure Active Directory → Properties → "Tenant ID"
2. **Client ID**: App Registration → Overview → "Application (client) ID"
3. **Client Secret**: De secret value die je in stap 2 hebt gekopieerd

## 📝 credentials.json Invullen

### Voorbeeld Voor Één Klant

```json
{
"LoginCredentials":[
    {
        "customername": "mrtn.blog",
        "ClientID": "12345678-1234-1234-1234-123456789012", 
        "Secret": "abcDEF123456~ghiJKL789012.mnoPQR345678",
        "TenantID": "87654321-4321-4321-4321-210987654321",
        "color": "#1f77b4"
    }
]
}
```

### Voorbeeld Voor Meerdere Klanten

```json
{
"LoginCredentials":[
    {
        "customername": "mrtn.blog",
        "ClientID": "12345678-1234-1234-1234-123456789012", 
        "Secret": "abcDEF123456~ghiJKL789012.mnoPQR345678",
        "TenantID": "87654321-4321-4321-4321-210987654321",
        "color": "#1f77b4"
    },
    {
        "customername": "Fabrikam", 
        "ClientID": "98765432-8765-4321-1234-567890123456",
        "Secret": "zyxWVU987654~tukRQP321098.lkjHGF654321", 
        "TenantID": "13579246-9753-1357-2468-135792468024",
        "color": "#ff7f0e"
    },
    {
        "customername": "Adventure Works",
        "ClientID": "11111111-2222-3333-4444-555555555555",
        "Secret": "qwerTY123456~asdfGH789012.zxcvBN345678",
        "TenantID": "66666666-7777-8888-9999-000000000000", 
        "color": "#2ca02c"
    }
]
}
```

## 🎨 Color Codes

De `color` property bepaalt de kleur in de HTML dashboard grafieken. Gebruik unieke hex kleuren per klant:

### Veelgebruikte Kleuren

| Kleur | Hex Code | Voorbeeld |
|-------|----------|-----------|
| Blauw | `#1f77b4` | 🔵 |
| Oranje | `#ff7f0e` | 🟠 |
| Groen | `#2ca02c` | 🟢 |
| Rood | `#d62728` | 🔴 |
| Paars | `#9467bd` | 🟣 |
| Bruin | `#8c564b` | 🤎 |
| Roze | `#e377c2` | 🩷 |
| Grijs | `#7f7f7f` | ⚫ |
| Geel | `#bcbd22` | 🟡 |
| Turquoise | `#17becf` | 🔷 |

### Online Color Picker

Voor custom kleuren kun je deze tools gebruiken:

- [Google Color Picker](https://g.co/kgs/color)
- [Adobe Color](https://color.adobe.com/)
- [Coolors.co](https://coolors.co/)

## ⚠️ Beveiliging & Best Practices

### 🔒 Secret Management

- **Nooit** secrets committen naar Git repositories
- Gebruik `.gitignore` om `credentials.json` uit te sluiten
- Roteer secrets regelmatig (elke 6-12 maanden)
- Gebruik descriptive names voor App Registrations per klant

### 📁 File Permissions

```powershell
# Zet restrictieve permissions op credentials file (Windows)
icacls "credentials.json" /inheritance:r /grant:r "$($env:USERNAME):(R,W)"
```

### 🔄 Secret Rotation

1. Genereer nieuwe secret in Azure Portal
2. Update `credentials.json` met nieuwe secret
3. Test de connectie
4. Verwijder oude secret in Azure Portal

## 🐛 Troubleshooting

### Veel Voorkomende Fouten

#### "Insufficient privileges to complete the operation"

**Oorzaak**: Admin consent niet gegeven voor API permissions  
**Oplossing**: Ga naar Azure Portal → App Registration → API permissions → "Grant admin consent"

#### "Application '[App-ID]' is not registered in directory '[Tenant-ID]'"

**Oorzaak**: Verkeerde Tenant ID of Client ID  
**Oplossing**: Controleer Tenant ID en Client ID in Azure Portal

#### "Invalid client secret is provided"

**Oorzaak**: Verkeerde of verlopen client secret  
**Oplossing**: Genereer nieuwe client secret en update credentials.json

#### "Forbidden - Insufficient privileges"

**Oorzaak**: Ontbrekende API permissions  
**Oplossing**: Controleer of alle required permissions zijn toegevoegd en admin consent is gegeven

### Verbinding Testen

Na het instellen kun je de verbinding testen door het script uit te voeren:

```powershell
.\get-windows-update-report.ps1
```

Het script zal per klant proberen verbinding te maken en eventuele fouten rapporteren.

## 📚 Gerelateerde Documentatie

- [README.md](README.md) - Hoofddocumentatie
- [CONFIG-UITLEG.md](CONFIG-UITLEG.md) - Configuratie opties
- [Microsoft Graph Permissions Reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Azure App Registration Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
