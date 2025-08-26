<#
.SYNOPSIS
Genereert een Windows Update rapportage voor meerdere tenants via Microsoft Graph.

.DESCRIPTION
Dit script haalt per tenant de ontbrekende Windows-updates op via de Microsoft Graph Threat Hunting API.
De resultaten worden geëxporteerd naar CSV-bestanden en een HTML-dashboard met filterbare tabellen en grafieken.

.BENODIGDHEDEN
- PowerShell 5+
- Microsoft Graph PowerShell SDK
- Een Azure AD App Registration per tenant met de juiste permissies

.GEBRUIK
1. Vul het credentials.json bestand met de juiste tenantgegevens.
2. config.json bevat de export instellingen zoals de export directory, archief directory en het aantal te behouden exports.
3. Voer het script uit: .\get-windows-update-report.ps1
4. Bekijk de resultaten in de map 'exports'.


.AUTEUR
Maarten Schmeitz (info@maarten-schmeitz.nl  | https://www.mrtn.blog)

.LASTEDIT
2025-08-13

.VERSIE
3.0
#>

#Versie-informatie
    $ProjectVersion = "3.0"
    $LastEditDate = "2025-08-13"

# Import configuratie
    try {
    $configJson = Get-Content -Path ".\config.json" -Raw
    $config = $configJson | ConvertFrom-Json
}
catch {
    Write-Error "Fout bij het laden of parsen van 'config.json': $($_.Exception.Message)"
    Write-Host "Controleer of 'config.json' aanwezig is, leesbaar is, en geldige JSON bevat." -ForegroundColor Red
    exit 1
}

# Functie voor het controleren en installeren van PowerShell modules
function Install-RequiredModules {
    param(
        [string[]]$ModuleNames
    )
    
    Write-Host "Controleren van benodigde PowerShell modules..." -ForegroundColor Cyan
    
    foreach ($ModuleName in $ModuleNames) {
        Write-Host "Verwerken van module: $ModuleName" -ForegroundColor White
        
        $Module = Get-Module -ListAvailable -Name $ModuleName
        
        if (-not $Module) {
            Write-Host "Module '$ModuleName' niet gevonden. Bezig met installeren..." -ForegroundColor Yellow
            try {
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
                Write-Host "Module '$ModuleName' succesvol geïnstalleerd." -ForegroundColor Green
            }
            catch {
                Write-Error "Fout bij installeren van module '$ModuleName': $($_.Exception.Message)"
                throw
            }
        }
        else {
            Write-Host "Module '$ModuleName' is al aanwezig." -ForegroundColor Green
        }
        
        # Importeer de module met expliciete feedback
        Write-Host "Importeren van module '$ModuleName'..." -ForegroundColor White
        try {
            # Probeer eerst alleen de benodigde commands te importeren
            Import-Module -Name $ModuleName -Force -ErrorAction Stop
            Write-Host "Module '$ModuleName' geïmporteerd." -ForegroundColor Green
        }
        catch {
            Write-Error "Fout bij importeren van module '$ModuleName': $($_.Exception.Message)"
            throw
        }
    }
    
    Write-Host "Module controle voltooid." -ForegroundColor Green
}

# Lijst van benodigde modules
$RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Security"
)

# Installeer en importeer benodigde modules
Install-RequiredModules -ModuleNames $RequiredModules

# Functie voor het controleren van App Registration geldigheid
function Test-AppRegistrationValidity {
    param(
        [string]$TenantID,
        [string]$ClientID,
        [System.Management.Automation.PSCredential]$ClientSecretCredential
    )
    
    try {
        # Verbind met Graph
        Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome | Out-Null
        
        # Haal app registration details op
        $App = Get-MgApplication -Filter "AppId eq '$ClientID'"
        
        if ($App -and $App.PasswordCredentials) {
            # Zoek naar de actieve client secret
            $ActiveSecret = $App.PasswordCredentials | Where-Object { 
                $_.EndDateTime -gt (Get-Date) 
            } | Sort-Object EndDateTime | Select-Object -First 1
            
            if ($ActiveSecret) {
                $ExpiryDate = $ActiveSecret.EndDateTime
                $DaysRemaining = [math]::Floor(($ExpiryDate - (Get-Date)).TotalDays)
                
                # Bepaal kleur op basis van dagen
                $Color = switch ($DaysRemaining) {
                    { $_ -gt 30 } { "Green" }
                    { $_ -ge 15 -and $_ -le 30 } { "Yellow" }
                    { $_ -lt 15 } { "Red" }
                    default { "Red" }
                }
                
                return @{
                    IsValid = $true
                    DaysRemaining = $DaysRemaining
                    ExpiryDate = $ExpiryDate
                    Color = $Color
                    Message = "$DaysRemaining dagen resterend"
                }
            }
        }
        
        return @{
            IsValid = $false
            DaysRemaining = 0
            ExpiryDate = $null
            Color = "Red"
            Message = "Geen geldige client secret gevonden"
        }
    }
    catch {
        return @{
            IsValid = $false
            DaysRemaining = 0
            ExpiryDate = $null
            Color = "Red"
            Message = "Fout bij controle: $($_.Exception.Message)"
        }
    }
    finally {
        try { Disconnect-MgGraph | Out-Null } catch { }
    }
}

# Helper functie om KB nummers en update identificaties te extraheren
function Get-CleanUpdateIdentifier {
    param(
        [string]$UpdateDisplayName
    )
    
    if (-not $UpdateDisplayName) {
        return ""
    }
    
    # Zoek naar KB nummers (KB gevolgd door cijfers)
    if ($UpdateDisplayName -match "(KB\d+)") {
        $KBNumber = $matches[1]
        
        # Zoek ook naar datum in formaat YYYY-MM
        if ($UpdateDisplayName -match "(\d{4}-\d{2})") {
            $DatePart = $matches[1]
            return "$DatePart Cumulative Update ($KBNumber)"
        } else {
            return "Cumulative Update ($KBNumber)"
        }
    }
    
    # Als er geen KB nummer is, probeer een verkorte versie van de naam
    if ($UpdateDisplayName -match "(\d{4}-\d{2}).*[Cc]umulative") {
        $DatePart = $matches[1]
        return "$DatePart Cumulative Update"
    } elseif ($UpdateDisplayName -match "[Cc]umulative") {
        return "Cumulative Update"
    } elseif ($UpdateDisplayName -match "Security.*Update") {
        return "Security Update"
    } elseif ($UpdateDisplayName -match "Feature.*Update") {
        return "Feature Update"
    } else {
        # Fallback: gebruik eerste 30 karakters van de naam
        $shortName = $UpdateDisplayName.Substring(0, [Math]::Min($UpdateDisplayName.Length, 30))
        if ($UpdateDisplayName.Length -gt 30) {
            $shortName += "..."
        }
        return $shortName
    }
}

# Helper functie om de nieuwste KB updates online op te halen
function Get-LatestKBUpdate {
    param(
        [string]$WindowsVersion,  # bijv. "Windows 10", "Windows 11"
        [string]$CurrentBuild,    # bijv. "4652"
        [string]$TargetBuild,     # bijv. "4946"
        [PSCustomObject]$Config   # Configuratie object met kbMappingUrl
    )
    
    try {
        # Bepaal Windows versie op basis van build nummer
        $WindowsProduct = "Windows 10"
        if ([int]$TargetBuild -ge 22000) {
            $WindowsProduct = "Windows 11"
        }
        
        # Online KB mapping URL uit config halen (met fallback)
        $OnlineKBUrl = if ($Config -and $Config.kbMappingUrl) { 
            $Config.kbMappingUrl 
        } else { 
            "https://mrtn.blog/wp-content/uploads/2025/08/kb-mapping.json" 
        }
        
        $KBNumber = $null
        $UpdateTitle = $null
        
        # Methode 1: Probeer online KB mapping op te halen
        try {
            $timeout = if ($Config.kbMapping.onlineTimeout) { $Config.kbMapping.onlineTimeout } else { 10 }
            Write-Verbose "Fetching KB mapping from: $OnlineKBUrl (timeout: $timeout seconds)"
            $OnlineMapping = Invoke-RestMethod -Uri $OnlineKBUrl -Method GET -TimeoutSec $timeout -ErrorAction Stop
            
            # Zoek in de juiste Windows versie sectie
            $MappingSection = if ([int]$TargetBuild -ge 22000) { 
                $OnlineMapping.mappings.windows11 
            } else { 
                $OnlineMapping.mappings.windows10 
            }
            
            # Zoek exacte match
            if ($MappingSection.$TargetBuild) {
                $FoundMapping = $MappingSection.$TargetBuild
                $KBNumber = $FoundMapping.kb
                # Check of de online title al "for Windows" bevat
                $baseTitle = $FoundMapping.title
                if ($baseTitle -notmatch "for Windows") {
                    $UpdateTitle = "$($FoundMapping.date) $baseTitle for $WindowsProduct ($KBNumber)"
                } else {
                    $UpdateTitle = "$($FoundMapping.date) $baseTitle ($KBNumber)"
                }
                Write-Verbose "Found exact online mapping: $UpdateTitle"
            } else {
                # Zoek dichtstbijzijnde build
                $ClosestBuild = $null
                $SmallestDifference = [int]::MaxValue
                
                foreach ($buildKey in $MappingSection.PSObject.Properties.Name) {
                    $difference = [Math]::Abs([int]$buildKey - [int]$TargetBuild)
                    if ($difference -lt $SmallestDifference) {
                        $SmallestDifference = $difference
                        $ClosestBuild = $buildKey
                    }
                }
                
                if ($ClosestBuild -and $SmallestDifference -lt 1000) {
                    $FoundMapping = $MappingSection.$ClosestBuild
                    $KBNumber = $FoundMapping.kb
                    # Check of de online title al "for Windows" bevat
                    $baseTitle = $FoundMapping.title
                    if ($baseTitle -notmatch "for Windows") {
                        $UpdateTitle = "$($FoundMapping.date) $baseTitle for $WindowsProduct ($KBNumber)"
                    } else {
                        $UpdateTitle = "$($FoundMapping.date) $baseTitle ($KBNumber)"
                    }
                    if ($SmallestDifference -gt 0 -and $Config.kbMapping.showEstimationLabels) {
                        $estimationLabel = if ($Config.kbMapping.estimationLabels.buildDifference) { 
                            $Config.kbMapping.estimationLabels.buildDifference -replace '\{targetBuild\}', $TargetBuild 
                        } else { 
                            "(geschat voor build $TargetBuild)" 
                        }
                        $UpdateTitle += " $estimationLabel"
                    }
                    Write-Verbose "Found closest online mapping: $UpdateTitle (difference: $SmallestDifference)"
                }
            }
        } catch {
            Write-Verbose "Failed to fetch online KB mapping: $($_.Exception.Message)"
        }
        
        # Methode 2: Fallback naar lokale KB mapping
        if (-not $KBNumber -and ($Config.kbMapping.fallbackToLocalMapping -ne $false)) {
            Write-Verbose "Using fallback local KB mapping"
            # Gebruik bekende patronen voor recente builds (bijgewerkt tot augustus 2025)
            $LocalKBMappings = @{
                # Windows 11 24H2 (2024-2025)
                "26100" = @{ KB = "KB5041585"; Date = "2024-08"; Title = "Cumulative Update" }
                "26010" = @{ KB = "KB5041580"; Date = "2024-08"; Title = "Cumulative Update" }
                
                # Windows 11 23H2 (2023-2024)
                "22631" = @{ KB = "KB5041585"; Date = "2024-08"; Title = "Cumulative Update" }
                "22621" = @{ KB = "KB5041592"; Date = "2024-08"; Title = "Cumulative Update" }
                
                # Windows 11 22H2 (2022-2023)
                "22000" = @{ KB = "KB5041580"; Date = "2024-08"; Title = "Cumulative Update" }
                
                # Windows 10 22H2 (2022-2025)
                "19045" = @{ KB = "KB5041580"; Date = "2024-08"; Title = "Cumulative Update" }
                "19044" = @{ KB = "KB5041580"; Date = "2024-08"; Title = "Cumulative Update" }
                "19043" = @{ KB = "KB5041577"; Date = "2024-08"; Title = "Cumulative Update" }
                "19042" = @{ KB = "KB5041577"; Date = "2024-08"; Title = "Cumulative Update" }
                "19041" = @{ KB = "KB5041568"; Date = "2024-08"; Title = "Cumulative Update" }
                
                # Windows 10 oudere versies (bijgewerkt naar recentere datums)
                "18363" = @{ KB = "KB5041561"; Date = "2024-08"; Title = "Cumulative Update" }
                "18362" = @{ KB = "KB5041561"; Date = "2024-07"; Title = "Cumulative Update" }
                "17763" = @{ KB = "KB5041564"; Date = "2024-08"; Title = "Cumulative Update" }
                "17134" = @{ KB = "KB5040442"; Date = "2024-07"; Title = "Cumulative Update" }
                "16299" = @{ KB = "KB5039338"; Date = "2024-06"; Title = "Cumulative Update" }
                "15063" = @{ KB = "KB5038223"; Date = "2024-05"; Title = "Cumulative Update" }
                "14393" = @{ KB = "KB5037019"; Date = "2024-04"; Title = "Cumulative Update" }
                "10586" = @{ KB = "KB5036004"; Date = "2024-03"; Title = "Cumulative Update" }
                "10240" = @{ KB = "KB5034768"; Date = "2024-08"; Title = "Cumulative Update" }
                
                # Fallback voor nieuwe builds (geschat)
                "27000" = @{ KB = "KB5042000"; Date = "2025-08"; Title = "Cumulative Update" }
                "26500" = @{ KB = "KB5041900"; Date = "2025-01"; Title = "Cumulative Update" }
            }
            
            # Zoek naar exacte match eerst
            if ($LocalKBMappings.ContainsKey($TargetBuild)) {
                $ClosestMapping = $LocalKBMappings[$TargetBuild]
            } else {
                # Zoek naar de dichtstbijzijnde mapping
                $ClosestMapping = $null
                $SmallestDifference = [int]::MaxValue
                
                foreach ($buildKey in $LocalKBMappings.Keys) {
                    $difference = [Math]::Abs([int]$buildKey - [int]$TargetBuild)
                    if ($difference -lt $SmallestDifference) {
                        $SmallestDifference = $difference
                        $ClosestMapping = $LocalKBMappings[$buildKey]
                    }
                }
            }
            
            if ($ClosestMapping) {
                $KBNumber = $ClosestMapping.KB
                # Check of de lokale title al "for Windows" bevat (dat zou niet moeten, maar ter zekerheid)
                $baseTitle = $ClosestMapping.Title
                if ($baseTitle -notmatch "for Windows") {
                    $UpdateTitle = "$($ClosestMapping.Date) $baseTitle for $WindowsProduct ($KBNumber)"
                } else {
                    $UpdateTitle = "$($ClosestMapping.Date) $baseTitle ($KBNumber)"
                }
                
                # Controleer of de mapping oud is (datum meer dan 6 maanden geleden)
                $mappingDate = try { [DateTime]::ParseExact($ClosestMapping.Date, "yyyy-MM", $null) } catch { $null }
                $isOldMapping = $mappingDate -and (Get-Date).AddMonths(-6) -gt $mappingDate
                
                # Bepaal de drempelwaarde voor geschatte mapping
                $threshold = if ($Config.kbMapping.estimationThreshold) { $Config.kbMapping.estimationThreshold } else { 1000 }
                
                # Voeg labels toe op basis van configuratie
                if ($Config.kbMapping.showEstimationLabels) {
                    if ($SmallestDifference -gt $threshold) {
                        $estimationLabel = if ($Config.kbMapping.estimationLabels.noMapping) { 
                            $Config.kbMapping.estimationLabels.noMapping 
                        } else { 
                            "(geschat)" 
                        }
                        $UpdateTitle += " $estimationLabel"
                    } elseif ($isOldMapping) {
                        $estimationLabel = if ($Config.kbMapping.estimationLabels.oldMapping) { 
                            $Config.kbMapping.estimationLabels.oldMapping 
                        } else { 
                            "(verouderd)" 
                        }
                        $UpdateTitle += " $estimationLabel"
                    }
                }
            }
        }
        
        # Methode 3: Genereer een geschatte KB op basis van datum en build verschil
        if (-not $KBNumber) {
            $CurrentDate = Get-Date
            $EstimatedMonth = $CurrentDate.ToString("yyyy-MM")
            $CurrentYear = $CurrentDate.Year
            
            # Genereer een realistische KB nummer gebaseerd op het patroon
            $KBPrefix = switch ($CurrentYear) {
                2025 { "506" }
                2024 { "504" }
                2023 { "503" }
                default { "504" }
            }
            
            $EstimatedKB = "KB$KBPrefix" + ([string]([int]$TargetBuild % 10000)).PadLeft(4, '0')
            $UpdateTitle = "$EstimatedMonth Cumulative Update for $WindowsProduct ($EstimatedKB)"
            
            # Voeg geschat label toe als configuratie dit aangeeft
            if ($Config.kbMapping.showEstimationLabels) {
                $estimationLabel = if ($Config.kbMapping.estimationLabels.noMapping) { 
                    $Config.kbMapping.estimationLabels.noMapping 
                } else { 
                    "(geschat)" 
                }
                $UpdateTitle += " $estimationLabel"
            }
        }
        
        return @{
            Success = $true
            KBNumber = $KBNumber
            UpdateTitle = $UpdateTitle
            WindowsProduct = $WindowsProduct
            Method = if ($OnlineMapping) { "Online" } elseif ($KBNumber) { "Local" } else { "Estimated" }
        }
        
    } catch {
        Write-Verbose "Error retrieving KB update information: $($_.Exception.Message)"
        return @{
            Success = $false
            KBNumber = $null
            UpdateTitle = "Cumulative Update vereist"
            WindowsProduct = "Windows"
            Method = "Fallback"
            Error = $_.Exception.Message
        }
    }
}

# Onderdruk Microsoft Graph statusberichten
$env:POWERSHELL_TELEMETRY_OPTOUT = "1"
$ProgressPreference = "SilentlyContinue"

# Controleer of de exports directory bestaat, zo niet: maak hem aan
$ExportDir = ".\$($config.exportDirectory)"
if (-not (Test-Path -Path $ExportDir -PathType Container)) {
    New-Item -Path $ExportDir -ItemType Directory | Out-Null
}

# Controleer of de archive directory bestaat, zo niet: maak hem aan
$ArchiveDir = ".\$($config.archiveDirectory)"
if (-not (Test-Path -Path $ArchiveDir -PathType Container)) {
    New-Item -Path $ArchiveDir -ItemType Directory | Out-Null
}

# Functie voor het verplaatsen van oude export bestanden naar archief
function Move-OldExportsToArchive {
    param(
        [string]$ExportPath,
        [string]$ArchivePath,
        [int]$RetentionCount
    )
    
    if ($config.cleanupOldExports -eq $true -and $RetentionCount -gt 0) {
        Write-Host "Bezig met archiveren van oude export bestanden... (behouden: $RetentionCount per type)" -ForegroundColor Cyan
        
        # Groepeer bestanden per type (Overview of ByUpdate) en per klant
        $AllFiles = Get-ChildItem -Path $ExportPath -Filter "*.csv" | Sort-Object Name -Descending
        
        # Groepeer per klant en type
        $GroupedFiles = $AllFiles | Group-Object { 
            # Verwacht patroon: Prefix_Customer_Type.csv
            $parts = $_.Name -split "_"
            # Controleer of het bestand voldoet aan het verwachte patroon
            if ($parts.Count -ge 3) {
                # Controleer of het laatste deel eindigt op .csv
                $typePart = $parts[-1]
                if ($typePart -match "^[A-Za-z]+\.csv$") {
                    return "$($parts[1])_$($typePart)" # CustomerName_Type.csv
                }
            }
            return "Unknown"
        }
        
        foreach ($group in $GroupedFiles) {
            $FilesToArchive = $group.Group | Select-Object -Skip $RetentionCount
            
            if ($FilesToArchive.Count -gt 0) {
                foreach ($file in $FilesToArchive) {
                    $DestinationPath = Join-Path -Path $ArchivePath -ChildPath $file.Name
                    Write-Host "Archiveren: $($file.Name) -> $ArchivePath" -ForegroundColor Yellow
                    Move-Item -Path $file.FullName -Destination $DestinationPath -Force
                }
                Write-Host "Groep '$($group.Name)': $($FilesToArchive.Count) bestanden gearchiveerd." -ForegroundColor Green
            }
        }
    }
}

#Import JSON
# Read the JSON file
$json = Get-Content -Path ".\credentials.json" -Raw

# Convert JSON to PowerShell object
$data = $json | ConvertFrom-Json

# Verzamel App Registration informatie
$AppRegistrationData = @{}

foreach ($cred in $data.LoginCredentials) {

    Write-Host "Verwerken van klant: $($cred.customername)" -ForegroundColor Cyan

    $ClientID = "$($cred.ClientID)"
    $Secret = "$($cred.Secret)"
    $TenantID = "$($cred.TenantID)"

    #Collect App Secret
    $Secret = ConvertTo-SecureString $Secret -AsPlainText -Force
    $ClientSecretCredential = New-Object System.Management.Automation.PSCredential -ArgumentList ($ClientID, $Secret)
    
    # Controleer App Registration geldigheid
    Write-Host "Controleren App Registration geldigheid..." -ForegroundColor White
    $AppValidity = Test-AppRegistrationValidity -TenantID $TenantID -ClientID $ClientID -ClientSecretCredential $ClientSecretCredential
    Write-Host "App Registration: $($AppValidity.Message)" -ForegroundColor $AppValidity.Color
    
    # Sla App Registration info op voor HTML rapport
    $AppRegistrationData[$cred.customername] = $AppValidity

    #Connect to Graph using Application Secret
    Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome | Out-Null

    Write-Host "Ophalen Windows Update status via Device Management..." -ForegroundColor Cyan

    try {
        # Probeer Windows devices op te halen via Device Management API
        $DevicesUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'"
        
        try {
            $DevicesResponse = Invoke-MgGraphRequest -Method GET -Uri $DevicesUri -ErrorAction Stop
            $Devices = $DevicesResponse.value
            
            Write-Host "Gevonden $($Devices.Count) Windows devices..." -ForegroundColor Green
        }
        catch {
            # Als Device Management API faalt (bijv. geen permissions), gebruik fallback
            Write-Warning "Device Management API niet beschikbaar voor deze tenant. Reden: $($_.Exception.Message)"
            Write-Host "Gebruik fallback methode met Threat Hunting API..." -ForegroundColor Yellow
            $Devices = @()  # Forceer fallback door lege array
        }
        
        if ($Devices.Count -eq 0) {
            Write-Warning "Geen Windows devices gevonden in Device Management. Probeer fallback met Threat Hunting..."
            
            # Fallback naar device informatie via KQL
            $FallbackQuery = "DeviceInfo 
            | where OSPlatform startswith 'Windows'
            | summarize arg_max(Timestamp, *) by DeviceName
            | project DeviceName, LastSeen=Timestamp, LoggedOnUsers, OSPlatform, OSVersion"
            
            $Body = @{ Query = $FallbackQuery } | ConvertTo-Json
            $FallbackResult = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/security/runHuntingQuery" -Body $Body
            
            $ResultsArray = $FallbackResult.results | ForEach-Object {
                [PSCustomObject]@{
                    DeviceName = $_.DeviceName
                    MissingUpdates = @("Windows Update status: Controleer handmatig - Device niet in Intune beheer")
                    Count = 1  # Handmatige controle vereist = niet up-to-date
                    LastSeen = $_.LastSeen
                    LoggedOnUsers = if ($_.LoggedOnUsers -is [System.Array]) { $_.LoggedOnUsers -join ', ' } else { $_.LoggedOnUsers }
                    OSPlatform = $_.OSPlatform
                    OSVersion = $_.OSVersion
                    UpdateStatus = "Handmatige controle vereist"
                }
            }
        } else {
            # Verwerk elk device voor Windows Update status
            $ResultsArray = @()
            
            foreach ($Device in $Devices) {
                try {
                    $DeviceName = $Device.deviceName
                    $MissingUpdates = @()
                    $ActualMissingUpdates = @()  # Voor echte KB nummers/update namen
                    $UpdateStatus = "Onbekend"
                    
                    # Controleer Windows Update compliance status
                    $ComplianceUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($Device.id)')/deviceCompliancePolicyStates"
                    try {
                        $ComplianceStates = Invoke-MgGraphRequest -Method GET -Uri $ComplianceUri -ErrorAction SilentlyContinue
                        
                        $HasUpdateIssues = $false
                        if ($ComplianceStates.value) {
                            foreach ($ComplianceState in $ComplianceStates.value) {
                                if ($ComplianceState.state -eq 'nonCompliant' -and 
                                    ($ComplianceState.settingName -like '*update*' -or 
                                     $ComplianceState.displayName -like '*Update*' -or
                                     $ComplianceState.displayName -like '*Windows*')) {
                                    $MissingUpdates += "Non-compliant: $($ComplianceState.displayName)"
                                    $HasUpdateIssues = $true
                                }
                            }
                        }
                        
                        if ($HasUpdateIssues) {
                            $UpdateStatus = "Compliance problemen"
                        }
                    } catch {
                        Write-Verbose "Compliance check failed for $DeviceName"
                    }
                    
                    # Controleer update installation geschiedenis via windowsUpdateStates
                    $UpdateStatesUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($Device.id)')/windowsUpdateStates"
                    try {
                        $UpdateStates = Invoke-MgGraphRequest -Method GET -Uri $UpdateStatesUri -ErrorAction SilentlyContinue
                        
                        if ($UpdateStates.value) {
                            foreach ($UpdateState in $UpdateStates.value) {
                                # Debug: Log alle update states om te zien wat beschikbaar is
                                Write-Verbose "Device: $DeviceName, Update: $($UpdateState.displayName), State: $($UpdateState.state), Quality: $($UpdateState.qualityUpdateClassification)"
                                
                                if ($UpdateState.state -eq 'pendingInstallation') {
                                    $MissingUpdates += "Pending: $($UpdateState.displayName)"
                                    $ActualMissingUpdates += Get-CleanUpdateIdentifier -UpdateDisplayName $UpdateState.displayName
                                    $UpdateStatus = "Updates wachtend"
                                } elseif ($UpdateState.state -eq 'failed') {
                                    $MissingUpdates += "Failed: $($UpdateState.displayName)"
                                    $ActualMissingUpdates += Get-CleanUpdateIdentifier -UpdateDisplayName $UpdateState.displayName
                                    $UpdateStatus = "Update fouten"
                                } elseif ($UpdateState.state -eq 'notApplicableForDevice') {
                                    # Skip deze - niet van toepassing
                                    continue
                                } elseif ($UpdateState.state -eq 'installed') {
                                    # Skip deze - al geïnstalleerd
                                    continue
                                } else {
                                    # Alle andere states kunnen duiden op missing/needed updates
                                    $MissingUpdates += "State $($UpdateState.state): $($UpdateState.displayName)"
                                    $ActualMissingUpdates += Get-CleanUpdateIdentifier -UpdateDisplayName $UpdateState.displayName
                                    if ($UpdateStatus -eq "Onbekend") {
                                        $UpdateStatus = "Updates beschikbaar"
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Verbose "Update states check failed for $DeviceName"
                    }
                    
                    # Alternatieve methode: probeer Windows Update for Business reports API
                    if ($ActualMissingUpdates.Count -eq 0) {
                        try {
                            # Check device configuration voor Windows Update settings
                            $ConfigUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($Device.id)')/deviceConfigurationStates"
                            $ConfigStates = Invoke-MgGraphRequest -Method GET -Uri $ConfigUri -ErrorAction SilentlyContinue
                            
                            if ($ConfigStates.value) {
                                foreach ($ConfigState in $ConfigStates.value) {
                                    if ($ConfigState.displayName -like "*Update*" -or $ConfigState.displayName -like "*Windows*") {
                                        if ($ConfigState.state -eq "nonCompliant") {
                                            $MissingUpdates += "Config non-compliant: $($ConfigState.displayName)"
                                            $ActualMissingUpdates += "Configuration Policy"
                                            if ($UpdateStatus -eq "Onbekend") {
                                                $UpdateStatus = "Configuratie problemen"
                                            }
                                        }
                                    }
                                }
                            }
                        } catch {
                            Write-Verbose "Config states check failed for $DeviceName"
                        }
                    }
                    
                    # Probeer aanvullende Windows Update informatie op te halen
                    if ($ActualMissingUpdates.Count -eq 0) {
                        try {
                            # Check voor Windows Update deployment states
                            $DeploymentUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($Device.id)')"
                            $DeviceDetails = Invoke-MgGraphRequest -Method GET -Uri $DeploymentUri -ErrorAction SilentlyContinue
                            
                            if ($DeviceDetails -and $DeviceDetails.windowsUpdateForBusinessConfiguration) {
                                Write-Verbose "Found Windows Update for Business config for $DeviceName"
                            }
                            
                        } catch {
                            Write-Verbose "Additional update check failed for $DeviceName"
                        }
                    }
                    
                    # Als geen specifieke problemen gevonden, bepaal status op basis van synchronisatie en OS versie
                    if ($MissingUpdates.Count -eq 0 -and $UpdateStatus -eq "Onbekend") {
                        $DaysSinceSync = if ($Device.lastSyncDateTime) {
                            [Math]::Round((New-TimeSpan -Start ([DateTime]$Device.lastSyncDateTime) -End (Get-Date)).TotalDays)
                        } else { 999 }
                        
                        # Controleer OS versie - voeg deze informatie toe voor transparantie
                        $OSVersion = if ($Device.osVersion) { $Device.osVersion } else { "Onbekend" }
                        
                        if ($DaysSinceSync -le 3) {
                            $MissingUpdates = @("Windows Update status: Recent gesynchroniseerd (OS: $OSVersion), geen update problemen gedetecteerd")
                            $UpdateStatus = "Up-to-date"
                            
                            # Voor machines die recent hebben gesynchroniseerd maar oudere OS hebben
                            if ($OSVersion -and $OSVersion -match '10\.0\.(\d+)\.(\d+)') {
                                $currentBuild = [int]$matches[2]
                                # Check voor bekende verouderde builds die updates nodig hebben
                                if ($currentBuild -lt 4946) {
                                    $ActualMissingUpdates += "Updates beschikbaar"
                                }
                            }
                        } elseif ($DaysSinceSync -le 7) {
                            $MissingUpdates = @("Windows Update status: Gesynchroniseerd binnen een week (OS: $OSVersion)")
                            $UpdateStatus = "Waarschijnlijk up-to-date"
                        } else {
                            $MissingUpdates = @("Windows Update status: Niet recent gesynchroniseerd ($DaysSinceSync dagen) (OS: $OSVersion)")
                            $UpdateStatus = "Synchronisatie vereist"
                        }
                    }
                    
                    $ResultsArray += [PSCustomObject]@{
                        DeviceName = $DeviceName
                        MissingUpdates = $MissingUpdates
                        ActualMissingUpdates = $ActualMissingUpdates
                        Count = if ($UpdateStatus -in @("Up-to-date", "Waarschijnlijk up-to-date")) { 0 } else { 1 }
                        LastSeen = $Device.lastSyncDateTime
                        LoggedOnUsers = if ($Device.userPrincipalName) { $Device.userPrincipalName } else { "Geen gebruiker" }
                        OSPlatform = $Device.operatingSystem
                        OSVersion = $Device.osVersion
                        UpdateStatus = $UpdateStatus
                    }
                    
                } catch {
                    Write-Warning "Fout bij verwerken device $($Device.deviceName): $($_.Exception.Message)"
                    
                    $ResultsArray += [PSCustomObject]@{
                        DeviceName = $Device.deviceName
                        MissingUpdates = @("Error: Kan Windows Update status niet controleren")
                        Count = 1  # Error = niet up-to-date
                        LastSeen = $Device.lastSyncDateTime
                        LoggedOnUsers = if ($Device.userPrincipalName) { $Device.userPrincipalName } else { "Geen gebruiker" }
                        OSPlatform = $Device.operatingSystem
                        OSVersion = if ($Device.osVersion) { $Device.osVersion } else { "Onbekend" }
                        UpdateStatus = "Error"
                    }
                }
            }
        }
        
        # Simuleer de oude Result structuur voor compatibiliteit
        $Result = @{
            results = $ResultsArray
        }
        
        # === OS VERSIE ANALYSE ===
        # Analyseer OS versies om machines met verouderde builds te identificeren
        $OSVersionGroups = $ResultsArray | Where-Object { $_.OSVersion -and $_.OSVersion -ne "Onbekend" } | 
                          Group-Object OSVersion | 
                          Sort-Object Name -Descending
        
        if ($OSVersionGroups.Count -gt 1) {
            $LatestOSVersion = $OSVersionGroups[0].Name
            $LatestCount = $OSVersionGroups[0].Count
            $TotalMachines = ($OSVersionGroups | Measure-Object Count -Sum).Sum
            
            Write-Host "`n=== OS VERSIE ANALYSE ===" -ForegroundColor Yellow
            Write-Host "Nieuwste OS versie gedetecteerd: $LatestOSVersion ($LatestCount machines)" -ForegroundColor Green
            Write-Host "Totaal machines: $TotalMachines" -ForegroundColor Cyan
            Write-Host "`nVerdeling per OS versie:" -ForegroundColor Cyan
            
            foreach ($group in $OSVersionGroups) {
                $percentage = [Math]::Round(($group.Count / $TotalMachines) * 100, 1)
                if ($group.Name -eq $LatestOSVersion) {
                    Write-Host "  [OK] $($group.Name): $($group.Count) machines ($percentage%)" -ForegroundColor Green
                } else {
                    Write-Host "  [WARNING] $($group.Name): $($group.Count) machines ($percentage%) - VEROUDERD" -ForegroundColor Yellow
                }
            }
            
            # Update de resultaten voor machines met verouderde OS versies
            $UpdatedResults = @()
            foreach ($result in $ResultsArray) {
                $newResult = $result.PSObject.Copy()
                
                if ($result.OSVersion -and $result.OSVersion -ne "Onbekend" -and $result.OSVersion -ne $LatestOSVersion) {
                    # Machines met verouderde OS versie - update hun status
                    $originalStatus = $result.UpdateStatus
                    $newResult.UpdateStatus = "Verouderde OS versie"
                    
                    # Voeg informatie toe over de verouderde OS versie
                    $versionDifference = "Huidige: $($result.OSVersion), Nieuwste: $LatestOSVersion"
                    $newResult.MissingUpdates += "LET OP: Verouderde OS versie - $versionDifference (was: $originalStatus)"
                    
                    # Probeer te bepalen welke cumulative update er nodig is op basis van OS versie
                    if ($result.OSVersion -and $LatestOSVersion) {
                        $currentBuild = $result.OSVersion -replace '.*\.(\d+)$', '$1'
                        $latestBuild = $LatestOSVersion -replace '.*\.(\d+)$', '$1'
                        
                        # Voor Windows 11/10 updates - haal KB nummers online op
                        if ($currentBuild -match '^\d+$' -and $latestBuild -match '^\d+$') {
                            $buildDifference = [int]$latestBuild - [int]$currentBuild
                            if ($buildDifference -gt 0) {
                                # Bepaal volledige build nummers voor Windows versie detectie
                                $fullCurrentBuild = $result.OSVersion -replace '^.*\.(\d+\.\d+)$', '$1'
                                $fullLatestBuild = $LatestOSVersion -replace '^.*\.(\d+\.\d+)$', '$1'
                                
                                # Extract major build (bijv. 26100 uit 26100.4652)
                                $majorCurrentBuild = if ($fullCurrentBuild -match '^(\d+)\.') { $matches[1] } else { $currentBuild }
                                $majorLatestBuild = if ($fullLatestBuild -match '^(\d+)\.') { $matches[1] } else { $latestBuild }
                                
                                # Haal de nieuwste KB update informatie online op
                                Write-Verbose "Looking up KB information for build $majorCurrentBuild -> $majorLatestBuild"
                                $KBInfo = Get-LatestKBUpdate -CurrentBuild $majorCurrentBuild -TargetBuild $majorLatestBuild -Config $config
                                
                                if ($KBInfo.Success -and $KBInfo.UpdateTitle) {
                                    $newResult.ActualMissingUpdates += $KBInfo.UpdateTitle
                                    Write-Verbose "Found KB info via $($KBInfo.Method): $($KBInfo.UpdateTitle)"
                                } elseif ($buildDifference -gt 100) {
                                    # Grote build verschillen duiden op multiple missing updates
                                    $newResult.ActualMissingUpdates += "Meerdere cumulative updates"
                                } elseif ($buildDifference -gt 50) {
                                    # Matige build verschillen
                                    $newResult.ActualMissingUpdates += "Cumulative update vereist"
                                } else {
                                    # Kleinere build verschillen - probeer generieke update info
                                    $CurrentDate = Get-Date
                                    $EstimatedMonth = $CurrentDate.ToString("yyyy-MM")
                                    $newResult.ActualMissingUpdates += "$EstimatedMonth Cumulative Update"
                                }
                            }
                        }
                    }
                    
                    $newResult.Count = 1  # Verouderde OS versie = niet up-to-date
                }
                
                $UpdatedResults += $newResult
            }
            
            # Update de result structure met de aangepaste resultaten
            $Result = @{
                results = $UpdatedResults
            }
            
            Write-Host "`nMachines met verouderde OS versies zijn gemarkeerd als 'Verouderde OS versie'" -ForegroundColor Yellow
            Write-Host "================================`n" -ForegroundColor Yellow
        } else {
            Write-Host "`n=== OS VERSIE ANALYSE ===" -ForegroundColor Green
            Write-Host "Alle machines hebben dezelfde OS versie: $($OSVersionGroups[0].Name)" -ForegroundColor Green
            Write-Host "============================`n" -ForegroundColor Green
        }
        
    } catch {
        Write-Warning "Fout bij ophalen Windows Update informatie voor $($cred.customername): $($_.Exception.Message)"
        Write-Host "Overslaan van deze klant en doorgaan met volgende..." -ForegroundColor Yellow
        
        # Maak lege results voor deze klant
        $Result = @{
            results = @([PSCustomObject]@{
                DeviceName = "Geen toegang"
                MissingUpdates = @("Error: Kan Windows Update informatie niet ophalen - mogelijk geen juiste permissions")
                ActualMissingUpdates = @()
                Count = 1  # Permission error = niet up-to-date
                LastSeen = (Get-Date).ToString()
                LoggedOnUsers = "N/A"
                OSPlatform = "Windows"
                OSVersion = "Onbekend" 
                UpdateStatus = "Permission Error"
            })
        }
    }

    #Format the results into an array
    $ResultsTable = $Result.results | ForEach-Object {
        $MissingUpdates = if ($_.MissingUpdates -is [System.Array]) { 
            $_.MissingUpdates -join ','
        } else { 
            $_.MissingUpdates 
        }
        
        $ActualMissingUpdates = if ($_.ActualMissingUpdates -is [System.Array]) { 
            $_.ActualMissingUpdates -join '; '
        } else { 
            $_.ActualMissingUpdates 
        }
        
        [PSCustomObject]@{
            Device = $_.DeviceName
            "Update Status" = if ($_.UpdateStatus) { $_.UpdateStatus } else { "Onbekend" }
            "Missing Updates" = if ($ActualMissingUpdates) { $ActualMissingUpdates } else { "" }
            "Details" = $MissingUpdates
            "OS Version" = if ($_.OSVersion) { $_.OSVersion } else { "Onbekend" }
            "Count" = if ($_.Count -gt 0) { $_.Count } else { 0 }
            "LastSeen" = $_.LastSeen
            "LoggedOnUsers" = if ($_.LoggedOnUsers -is [System.Array]) { $_.LoggedOnUsers -join ', ' } else { $_.LoggedOnUsers }
            "App Registration Days Remaining" = $AppValidity.DaysRemaining
            "App Registration Status" = $AppValidity.Message
        }
    }

    #Export the results
    $DateStamp = Get-Date -Format "yyyyMMdd"
    $ExportPath = ".\$($config.exportDirectory)\$DateStamp`_$($cred.customername)_Windows_Update_report_Overview.csv"
    $ResultsTable | Export-Csv -NoTypeInformation -path $ExportPath

    # Create a table: Missing Update -> Devices
    $MissingUpdateTable = @()
    foreach ($row in $Result.results) {
        # Use actual missing updates if available, otherwise fall back to details
        $UpdatesToProcess = if ($row.ActualMissingUpdates -and $row.ActualMissingUpdates.Count -gt 0) {
            $row.ActualMissingUpdates
        } else {
            $row.Details
        }
        
        foreach ($update in $UpdatesToProcess) {
            $MissingUpdateTable += [PSCustomObject]@{
                "Missing Update" = $update
                "Device" = $row.DeviceName
                "LastSeen" = $row.LastSeen
                "LoggedOnUsers" = if ($row.LoggedOnUsers -is [System.Array]) { $row.LoggedOnUsers -join ', ' } else { $row.LoggedOnUsers }
            }
        }
    }

    # Export Missing Update -> Devices table
    $ExportPathUpdates = ".\$($config.exportDirectory)\$DateStamp`_$($cred.customername)_Windows_Update_report_ByUpdate.csv"
    $MissingUpdateTable | Export-Csv -NoTypeInformation -Path $ExportPathUpdates

    #Disconnect from MG Graph
    Disconnect-MgGraph | Out-Null
    
    Write-Host "Klant $($cred.customername) voltooid. CSV-bestanden gegenereerd. App Registration: $($AppValidity.Message)" -ForegroundColor Green
}

# Voer archivering uit na alle exports
Move-OldExportsToArchive -ExportPath $ExportDir -ArchivePath $ArchiveDir -RetentionCount $config.exportRetentionCount

# Verzamel alle Overview-bestanden
$OverviewFiles = Get-ChildItem -Path "$ExportDir" -Filter "*_Overview.csv" | Sort-Object Name

# Haal de totalen per dag per klant op
$CountsPerDayPerCustomer = @{}
$ClientsPerDayPerCustomer = @{}
$LatestDatePerCustomer = @{}
$LatestCsvPerCustomer = @{}
foreach ($file in $OverviewFiles) {
    $csv = Import-Csv $file.FullName
    $parts = $file.Name -split "_"
    $Date = $parts[0]
    $Customer = $parts[1]
    $now = Get-Date
    $filterDays = [int]$config.lastSeenDaysFilter
    $filteredRows = @()
    foreach ($row in $csv) {
        $includeRow = $true
        if ($filterDays -gt 0 -and $row.LastSeen) {
            $lastSeenDate = $null
            try {
                $lastSeenDate = [datetime]::Parse($row.LastSeen)
            } catch {}
            if ($lastSeenDate) {
                $daysAgo = ($now - $lastSeenDate).TotalDays
                if ($daysAgo -gt $filterDays) {
                    $includeRow = $false
                }
            }
        }
        if ($includeRow) {
            $filteredRows += $row
        }
    }
    $TotalCount = ($filteredRows | Measure-Object -Property Count -Sum).Sum
    $ClientCount = $filteredRows.Count
    if (-not $CountsPerDayPerCustomer.ContainsKey($Customer)) {
        $CountsPerDayPerCustomer[$Customer] = @()
    }
    if (-not $ClientsPerDayPerCustomer.ContainsKey($Customer)) {
        $ClientsPerDayPerCustomer[$Customer] = @()
    }
    $CountsPerDayPerCustomer[$Customer] += [PSCustomObject]@{
        Date = $Date
        TotalCount = $TotalCount
    }
    $ClientsPerDayPerCustomer[$Customer] += [PSCustomObject]@{
        Date = $Date
        ClientCount = $ClientCount
    }
    # Bepaal de laatste datum per klant
    if (-not $LatestDatePerCustomer.ContainsKey($Customer) -or ($Date -gt $LatestDatePerCustomer[$Customer])) {
        $LatestDatePerCustomer[$Customer] = $Date
        $LatestCsvPerCustomer[$Customer] = $filteredRows
    }
}

# Verzamel alle unieke datums en sorteer ze
$AllDates = @()
foreach ($Customer in $CountsPerDayPerCustomer.Keys) {
    foreach ($DataPoint in $CountsPerDayPerCustomer[$Customer]) {
        if ($AllDates -notcontains $DataPoint.Date) {
            $AllDates += $DataPoint.Date
        }
    }
}
$AllDates = $AllDates | Sort-Object
$ChartLabels = $AllDates | ForEach-Object { "'$_'" }
$ChartLabelsString = $ChartLabels -join ","

# Genereer Chart.js datasets per klant (2 lijnen per klant)
$ChartDatasets = ""
$ChartDataJSON = "{"
foreach ($Customer in ($CountsPerDayPerCustomer.Keys | Sort-Object)) {
    # Maak hashtables voor snelle lookup van data per datum
    $CustomerCountLookup = @{}
    $CustomerClientLookup = @{}
    $now = Get-Date
    $filterDays = [int]$config.lastSeenDaysFilter
    foreach ($DataPoint in $CountsPerDayPerCustomer[$Customer]) {
        $includeRow = $true
        if ($filterDays -gt 0 -and $DataPoint.PSObject.Properties["LastSeen"] -and $DataPoint.LastSeen) {
            $lastSeenDate = $null
            try {
                $lastSeenDate = [datetime]::Parse($DataPoint.LastSeen)
            } catch {}
            if ($lastSeenDate) {
                $daysAgo = ($now - $lastSeenDate).TotalDays
                if ($daysAgo -gt $filterDays) {
                    $includeRow = $false
                }
            }
        }
        if ($includeRow) {
            $CustomerCountLookup[$DataPoint.Date] = $DataPoint.TotalCount
        }
    }
    foreach ($DataPoint in $ClientsPerDayPerCustomer[$Customer]) {
        $includeRow = $true
        if ($filterDays -gt 0 -and $DataPoint.PSObject.Properties["LastSeen"] -and $DataPoint.LastSeen) {
            $lastSeenDate = $null
            try {
                $lastSeenDate = [datetime]::Parse($DataPoint.LastSeen)
            } catch {}
            if ($lastSeenDate) {
                $daysAgo = ($now - $lastSeenDate).TotalDays
                if ($daysAgo -gt $filterDays) {
                    $includeRow = $false
                }
            }
        }
        if ($includeRow) {
            $CustomerClientLookup[$DataPoint.Date] = $DataPoint.ClientCount
        }
    }
    # Bouw data arrays met null-waarden voor ontbrekende datums
    $CountDataArray = @()
    $ClientDataArray = @()
    foreach ($Date in $AllDates) {
        if ($CustomerCountLookup.ContainsKey($Date)) {
            $CountDataArray += $CustomerCountLookup[$Date]
        } else {
            $CountDataArray += "null"
        }
        if ($CustomerClientLookup.ContainsKey($Date)) {
            $ClientDataArray += $CustomerClientLookup[$Date]
        } else {
            $ClientDataArray += "null"
        }
    }
    $CountData = $CountDataArray -join ","
    $ClientData = $ClientDataArray -join ","
    
    $credObj = $data.LoginCredentials | Where-Object { $_.customername -eq $Customer }
    $HexColor = $credObj.color
    # Fallback: als geen kleur, gebruik blauw
    if (-not $HexColor) { $HexColor = '#1f77b4' }

    # Dataset 1: Updates (volle lijn)
    $ChartDatasets += @"
        {
            label: '$Customer - Updates',
            data: [$CountData],
            borderColor: '$HexColor',
            backgroundColor: '$HexColor',
            fill: false,
            tension: 0.2,
            spanGaps: true,
            borderWidth: 2
        },
"@

    # Dataset 2: Clients (gestippelde lijn)
    $ChartDatasets += @"
        {
            label: '$Customer - Clients',
            data: [$ClientData],
            borderColor: '$HexColor',
            backgroundColor: '$HexColor',
            fill: false,
            tension: 0.2,
            spanGaps: true,
            borderWidth: 2,
            borderDash: [5, 5]
        },
"@

    # Voeg data toe voor individuele klant grafieken
    $CustomerLabels = ($CountsPerDayPerCustomer[$Customer] | ForEach-Object { "'$($_.Date)'" })
    $CustomerCountData = ($CountsPerDayPerCustomer[$Customer] | ForEach-Object { $_.TotalCount }) -join ","
    $CustomerClientData = ($ClientsPerDayPerCustomer[$Customer] | ForEach-Object { $_.ClientCount }) -join ","

    $ChartDataJSON += @"
    '$Customer': {
        labels: [$($CustomerLabels -join ",")],
        countData: [$CustomerCountData],
        clientData: [$CustomerClientData],
        borderColor: '$HexColor',
        backgroundColor: '$HexColor'
    },
"@
}

$ChartDataJSON = $ChartDataJSON.TrimEnd(',') + "}"

# Genereer tabbladen en tabellen voor alleen de laatste datum per klant
$CustomerTabs = ""
$CustomerTables = ""
foreach ($Customer in ($LatestCsvPerCustomer.Keys | Sort-Object)) {
    $TableRows = ""
    $RowCount = 0
    $now = Get-Date
    $filterDays = [int]$config.lastSeenDaysFilter
    foreach ($row in $LatestCsvPerCustomer[$Customer]) {
        $includeRow = $true
        if ($filterDays -gt 0 -and $row.LastSeen) {
            $lastSeenDate = $null
            try {
                $lastSeenDate = [datetime]::Parse($row.LastSeen)
            } catch {}
            if ($lastSeenDate) {
                $daysAgo = ($now - $lastSeenDate).TotalDays
                if ($daysAgo -gt $filterDays) {
                    $includeRow = $false
                }
            }
        }
        if ($includeRow) {
            # Voeg status kleuren toe
            $StatusColor = switch ($row.'Update Status') {
                "Up-to-date" { "color: green; font-weight: bold;" }
                "Waarschijnlijk up-to-date" { "color: darkgreen;" }
                "Verouderde OS versie" { "color: #FF8C00; font-weight: bold;" }
                "Compliance problemen" { "color: red; font-weight: bold;" }
                "Updates wachtend" { "color: orange; font-weight: bold;" }
                "Update fouten" { "color: red; font-weight: bold;" }
                "Synchronisatie vereist" { "color: orange;" }
                "Error" { "color: red; font-weight: bold;" }
                "Handmatige controle vereist" { "color: gray;" }
                default { "color: black;" }
            }
            
            $TableRows += "<tr><td>$($row.Device)</td><td style='$StatusColor'>$($row.'Update Status')</td><td>$($row.'Missing Updates')</td><td>$($row.'OS Version')</td><td>$($row.Count)</td><td>$($row.LastSeen)</td><td>$($row.LoggedOnUsers)</td></tr>`n"
            $RowCount++
        }
    }
    $CustomerTabs += '<button class="tablinks" onclick="openCustomer(event, ''' + $Customer + ''')">' + $Customer + ' (' + $RowCount + ')</button>'

    if ($filterDays -eq 0) {
       $lastSeenText = 'Dit zijn alle machines die gevonden kunnen worden.'
    } else {
        $lastSeenText = "Deze machines zijn de laatste $filterDays dagen online geweest."
    }
    $CustomerTables += @"
    <div id="$Customer" class="tabcontent" style="display:none">
        <h2>Laatste overzicht voor $Customer ($($LatestDatePerCustomer[$Customer]))</h2>
        <p style='font-style:italic;color:#555;'>$lastSeenText</p>
        
        <!-- Snelfilter knoppen voor Update Status -->
        <div class="filter-container">
            <strong>Snelfilters Update Status:</strong><br>
            <button onclick="filterByStatus('overviewTable_$Customer', '')" class="filter-button" style="background-color: #6c757d; color: white;">Alle statussen</button>
            <button onclick="filterByStatus('overviewTable_$Customer', 'Up-to-date')" class="filter-button" style="background-color: #28a745; color: white;">Up-to-date</button>
            <button onclick="filterByStatus('overviewTable_$Customer', 'Verouderde OS versie')" class="filter-button" style="background-color: #fd7e14; color: white;">Verouderde OS</button>
            <button onclick="filterByStatus('overviewTable_$Customer', 'Handmatige controle vereist')" class="filter-button" style="background-color: #dc3545; color: white;">Handmatige controle</button>
            <button onclick="filterByStatus('overviewTable_$Customer', 'Waarschijnlijk up-to-date')" class="filter-button" style="background-color: #17a2b8; color: white;">Waarschijnlijk up-to-date</button>
            <button onclick="filterByStatus('overviewTable_$Customer', 'Error')" class="filter-button" style="background-color: #6f42c1; color: white;">Errors</button>
        </div>
        
        <button onclick="exportTableToCSV('overviewTable_$Customer', '$Customer-full.csv', false)">Exporteren volledige tabel</button>
        <button onclick="exportTableToCSV('overviewTable_$Customer', '$Customer-filtered.csv', true)">Exporteren gefilterde rijen</button>
        <table id="overviewTable_$Customer" class="display" style="width:100%">
            <thead>
                <tr>
                    <th>Device</th>
                    <th>Update Status</th>
                    <th>Missing Updates</th>
                    <th>OS Version</th>
                    <th>Count</th>
                    <th>LastSeen</th>
                    <th>LoggedOnUsers</th>
                </tr>
            </thead>
            <tbody>
                $TableRows
            </tbody>
        </table>
    </div>
"@
}

$DataTablesScript = @'
$(document).ready(function() {
    // Initialiseer DataTables pas als een tab wordt geopend
    function initializeDataTable(tableId) {
        if (!$.fn.DataTable.isDataTable('#' + tableId)) {
            var table = $('#' + tableId).DataTable({
                "order": [[2, "desc"]],
                "language": {
                    "url": "//cdn.datatables.net/plug-ins/1.13.6/i18n/nl-NL.json"
                },
                "lengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]]
            });
            // Kolomfilters toevoegen
            $('#' + tableId + ' thead th').each(function (i) {
                var title = $(this).text();
                
                // Voor "Update Status" kolom (index 1): gebruik dropdown filter
                if (i === 1 && title.includes('Update Status')) {
                    // Verzamel unieke waarden uit de kolom
                    var uniqueValues = [];
                    table.column(i).data().unique().sort().each(function (d, j) {
                        if (d && uniqueValues.indexOf(d) === -1) {
                            uniqueValues.push(d);
                        }
                    });
                    
                    // Maak dropdown met unieke waarden
                    var select = '<br><select style="width:90%;font-size:12px;"><option value="">Alle statussen</option>';
                    uniqueValues.forEach(function(value) {
                        select += '<option value="' + value + '">' + value + '</option>';
                    });
                    select += '</select>';
                    
                    $(this).append(select);
                    $(this).find("select").on('change', function () {
                        var val = $.fn.dataTable.util.escapeRegex($(this).val());
                        table.column(i).search(val ? '^' + val + '$' : '', true, false).draw();
                    });
                } else {
                    // Voor andere kolommen: gebruik tekstfilter
                    $(this).append('<br><input type="text" placeholder="Filter '+title+'" style="width:90%;font-size:12px;" />');
                    $(this).find("input").on('keyup change', function () {
                        if (table.column(i).search() !== this.value) {
                            table.column(i).search(this.value).draw();
                        }
                    });
                }
            });
        }
    }
    // Maak initializeDataTable globaal beschikbaar
    window.initializeDataTable = initializeDataTable;

    // Export functie voor CSV
    window.exportTableToCSV = function(tableId, filename) {
        var csv = [];
        var table = $('#' + tableId).DataTable();
        // Header
        var header = [];
        $('#' + tableId + ' thead th').each(function() {
            header.push('"' + $(this).text().replace(/"/g, '""') + '"');
        });
        csv.push(header.join(','));
        // Data rows
        var onlyFiltered = arguments.length > 2 ? arguments[2] : false;
        var rowsToExport = onlyFiltered ? table.rows({ search: 'applied' }) : table.rows();
        rowsToExport.every(function(rowIdx, tableLoop, rowLoop) {
            var data = this.data();
            if (Array.isArray(data)) {
                var row = data.map(function(cell) {
                    return '"' + String(cell).replace(/"/g, '""') + '"';
                });
                csv.push(row.join(','));
            } else {
                var row = [];
                $(this.node()).find('td').each(function() {
                    row.push('"' + $(this).text().replace(/"/g, '""') + '"');
                });
                csv.push(row.join(','));
            }
        });
        var csvString = csv.join('\n');
        var blob = new Blob([csvString], { type: 'text/csv' });
        var link = document.createElement('a');
        link.href = window.URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
    
    // Functie voor snelfilters op Update Status
    window.filterByStatus = function(tableId, status) {
        var table = $('#' + tableId).DataTable();
        // Update Status kolom is index 1
        if (status === '') {
            // Reset filter - toon alle statussen
            table.column(1).search('').draw();
            // Reset ook de dropdown
            $('#' + tableId + ' thead th:eq(1) select').val('');
        } else {
            // Filter op specifieke status
            var val = $.fn.dataTable.util.escapeRegex(status);
            table.column(1).search('^' + val + '$', true, false).draw();
            // Update ook de dropdown
            $('#' + tableId + ' thead th:eq(1) select').val(status);
        }
    }
});
'@

$LastRunDate = Get-Date -Format "dd-MM-yyyy"

$Html = @"
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <title>Windows Update Overview</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css"/>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css"/>
    <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script>
    $DataTablesScript
    // Dark mode toggle with manual config only
    document.addEventListener('DOMContentLoaded', function() {
        var configThemeDefault = "$($config.theme.default)".toLowerCase();
        var btn = document.getElementById('darkModeToggle');
        function setTheme(useDark) {
            if (useDark) {
                document.body.classList.add('darkmode');
                if (btn) btn.innerHTML = '<i class="fa-solid fa-sun"></i> Light mode';
            } else {
                document.body.classList.remove('darkmode');
                if (btn) btn.innerHTML = '<i class="fa-solid fa-moon"></i> Dark mode';
            }
        }
        // Set initial theme and button state
        setTheme(configThemeDefault === "dark");
        if (btn) {
            btn.addEventListener('click', function() {
                var isDark = document.body.classList.contains('darkmode');
                setTheme(!isDark);
            });
        }
    });
    </script>
    <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #fff; color: #222; transition: background 0.2s, color 0.2s; }
    .container { max-width: 1200px; margin: auto; }
    canvas { background: #fff; }
    table.dataTable thead th { background: #eee; }
    .tab { overflow: hidden; border-bottom: 1px solid #ccc; }
    .tab button { background-color: #f1f1f1; float: left; border: none; outline: none; cursor: pointer; padding: 10px 20px; transition: 0.3s; }
    .tab button:hover { background-color: #ddd; }
    .tab button.active { background-color: #ccc; }
    .tabcontent { display: none; padding: 20px 0; }
    .footer { margin-top: 40px; padding: 20px 0; border-top: 1px solid #ddd; text-align: center; color: #666; font-size: 14px; }
    .footer a { color: #0066cc; text-decoration: none; }
    .footer a:hover { text-decoration: underline; }
    /* Dark mode styles */
    body.darkmode { background: #181a1b; color: #eee; }
    body.darkmode .container { background: #181a1b; }
    body.darkmode canvas { background: #222; }
    body.darkmode table.dataTable thead th { background: #222; color: #eee; }
    body.darkmode .tab { border-bottom: 1px solid #444; }
    body.darkmode .tab button { background-color: #222; color: #eee; }
    body.darkmode .tab button:hover { background-color: #333; }
    body.darkmode .tab button.active { background-color: #444; }
    body.darkmode .tabcontent { background: #181a1b; color: #eee; }
    body.darkmode .footer { border-top: 1px solid #444; color: #aaa; }
    body.darkmode .footer a { color: #66aaff; }
    
    /* Snelfilter knoppen styling */
    .filter-container { margin-bottom: 15px; padding: 10px; background-color: #f8f9fa; border-radius: 5px; }
    body.darkmode .filter-container { background-color: #2d3035; }
    .filter-button { margin: 2px; padding: 5px 10px; border: none; border-radius: 3px; cursor: pointer; font-size: 12px; }
    .filter-button:hover { opacity: 0.8; }
    </style>
</head>
<body>
<div class="container">
    <h1>Windows Update Overview <button id="darkModeToggle" style="float:right;margin-left:20px;"><i class="fa-solid fa-moon"></i> Dark mode</button></h1>

    <p>Laatst uitgevoerd op: $LastRunDate</p>
    <h2>Totale Count per dag per klant</h2>
    <canvas id="countChart" height="100"></canvas>
    <div class="tab">
        <button class="tablinks" onclick="showAllCustomers()">Alle klanten</button>
        <button class="tablinks" onclick="showAppRegistrations()">App Registrations</button>
        $CustomerTabs
    </div>
    
    <!-- App Registrations Tab -->
    <div id="AppRegistrations" class="tabcontent" style="display:none">
        <h2>App Registration Status Overview</h2>
        <table id="appRegTable" class="display" style="width:100%">
            <thead>
                <tr>
                    <th>Customer</th>
                    <th>Status</th>
                    <th>Days Remaining</th>
                    <th>Expiry Date</th>
                </tr>
            </thead>
            <tbody>
$(foreach ($customer in $AppRegistrationData.Keys | Sort-Object) {
    $appInfo = $AppRegistrationData[$customer]
    $statusColor = switch ($appInfo.Color) {
        "Green" { "color:green;" }
        "Yellow" { "color:orange;" }
        "Red" { "color:red;" }
        default { "color:gray;" }
    }
    $expiryDateFormatted = if ($appInfo.ExpiryDate) { $appInfo.ExpiryDate.ToString("dd-MM-yyyy") } else { "N/A" }
    "<tr><td>$customer</td><td style='$statusColor'>$($appInfo.Message)</td><td>$($appInfo.DaysRemaining)</td><td>$expiryDateFormatted</td></tr>"
} -join "`n")
            </tbody>
        </table>
    </div>
    
    $CustomerTables
    
    <div class="footer">
    Powered by <a href="https://github.com/scns/Windows-Update-Report-MultiTenant" target="_blank">Windows Update MultiTenant</a> <span style="font-weight:normal;color:#888;">v$ProjectVersion $LastEditDate</span> by  <a href="https://mrtn.blog" target="_blank">Maarten Schmeitz (mrtn.blog)</a>
    </div>
</div>
<script>
    // Chart.js data per klant
    const customerChartData = $ChartDataJSON;
    let chart;
    
    // Initialiseer de grafiek
    const ctx = document.getElementById('countChart').getContext('2d');
    chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: [$ChartLabelsString],
            datasets: [
                $ChartDatasets
            ]
        },
        options: {
            responsive: true,
            plugins: { 
                legend: { display: true },
                title: {
                    display: true,
                    text: 'Alle klanten'
                }
            },
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });

    // Tabs functie - nu met grafiek filtering
    function openCustomer(evt, customerName) {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        document.getElementById(customerName).style.display = "block";
        evt.currentTarget.className += " active";
        
        // Initialiseer DataTable voor deze klant
        if (typeof initializeDataTable === 'function') {
            initializeDataTable('overviewTable_' + customerName);
        }
        
        // Update grafiek voor specifieke klant
        if (customerChartData[customerName]) {
            chart.data.labels = customerChartData[customerName].labels;
            chart.data.datasets = [{
                label: customerName + ' - Updates',
                data: customerChartData[customerName].countData,
                borderColor: customerChartData[customerName].borderColor,
                backgroundColor: customerChartData[customerName].backgroundColor,
                fill: false,
                tension: 0.2,
                borderWidth: 2
            }, {
                label: customerName + ' - Clients',
                data: customerChartData[customerName].clientData,
                borderColor: customerChartData[customerName].borderColor.replace('rgb(', 'rgba(').replace(')', ', 0.6)'),
                backgroundColor: customerChartData[customerName].backgroundColor.replace('rgb(', 'rgba(').replace(')', ', 0.6)'),
                fill: false,
                tension: 0.2,
                borderWidth: 2,
                borderDash: [5, 5]
            }];
            chart.options.plugins.title.text = customerName;
            chart.update();
        }
    }
    
    // Functie voor App Registrations tab
    function showAppRegistrations() {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        document.getElementById("AppRegistrations").style.display = "block";
        document.getElementsByClassName("tablinks")[1].className += " active";
        
        // Initialiseer DataTable voor App Registrations
        if (typeof initializeDataTable === 'function') {
            initializeDataTable('appRegTable');
        }
        
        // Reset grafiek naar alle klanten zonder het tab te veranderen
        chart.data.labels = [$ChartLabelsString];
        chart.data.datasets = [
            $ChartDatasets
        ];
        chart.options.plugins.title.text = 'Alle klanten';
        chart.update();
    }

    // Functie om alle klanten te tonen
    function showAllCustomers() {
        // Verberg alle tabcontent
        var tabcontent = document.getElementsByClassName("tabcontent");
        for (var i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        
        // Verwijder active class van alle tabs
        var tablinks = document.getElementsByClassName("tablinks");
        for (var i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        
        // Activeer "Alle klanten" tab
        document.getElementsByClassName("tablinks")[0].className += " active";
        
        // Reset grafiek naar alle klanten
        chart.data.labels = [$ChartLabelsString];
        chart.data.datasets = [
            $ChartDatasets
        ];
        chart.options.plugins.title.text = 'Alle klanten';
        chart.update();
    }
    
    // Initialiseer de pagina bij het laden
    window.onload = function() {
        showAllCustomers();
    }
</script>
</body>
</html>
"@

$HtmlPath = ".\$($config.exportDirectory)\Windows_Update_Overview.html"
Set-Content -Path $HtmlPath -Value $Html -Encoding UTF8

# Open het HTML bestand automatisch in de standaard webbrowser (indien geconfigureerd)
Write-Host "HTML rapport gegenereerd: $HtmlPath" -ForegroundColor Green

if ($config.autoOpenHtmlReport -eq $true) {
    Write-Host "Openen van rapport in standaard webbrowser..." -ForegroundColor Cyan
    try {
        if ((Test-Path $HtmlPath -PathType Leaf) -and ($HtmlPath.ToLower().EndsWith(".html")))
        {
            Start-Process $HtmlPath
            Write-Host "Rapport succesvol geopend in webbrowser." -ForegroundColor Green
        }
        else {
            Write-Warning "Het HTML rapportbestand bestaat niet: $HtmlPath"
            Write-Host "U kunt het rapport handmatig openen via: $HtmlPath" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "Kon het rapport niet automatisch openen: $($_.Exception.Message)"
        Write-Host "U kunt het rapport handmatig openen via: $HtmlPath" -ForegroundColor Yellow
    }
}
else {
    Write-Host "Automatisch openen van rapport is uitgeschakeld in configuratie." -ForegroundColor Yellow
    Write-Host "U kunt het rapport handmatig openen via: $HtmlPath" -ForegroundColor Cyan
}

Write-Host "`nScript voltooid! Alle rapporten zijn gegenereerd en beschikbaar in de exports directory." -ForegroundColor Green
###############################################################
# BACKUP LOGIC: Zip exports, archive, config/credentials
###############################################################


# Backup root and subfolders from config
$BackupRoot = "./$($config.backup.backupRoot)"
$BackupExportDir = Join-Path $BackupRoot $config.backup.exportBackupSubfolder
$BackupArchiveDir = Join-Path $BackupRoot $config.backup.archiveBackupSubfolder
$BackupConfigDir = Join-Path $BackupRoot $config.backup.configBackupSubfolder

# Ensure backup folders exist
foreach ($dir in @($BackupRoot, $BackupExportDir, $BackupArchiveDir, $BackupConfigDir)) {
    if (-not (Test-Path -Path $dir -PathType Container)) {
        New-Item -Path $dir -ItemType Directory | Out-Null
    }
}

# Helper: Create zip and enforce retention
function New-ZipBackup {
    param(
        [string]$SourcePath,
        [string]$BackupFolder,
        [string]$BackupPrefix,
        [int]$RetentionCount
    )
    $timestamp = Get-Date -Format "yyyyMMdd"
    $zipName = "$BackupPrefix-$timestamp.zip"
    $zipPath = Join-Path $BackupFolder $zipName
    Compress-Archive -Path $SourcePath -DestinationPath $zipPath -Force
    # Retention: Remove oldest if over limit
    $zips = Get-ChildItem -Path $BackupFolder -Filter "$BackupPrefix-*.zip" | Sort-Object LastWriteTime -Descending
    if ($zips.Count -gt $RetentionCount) {
        $zipsToRemove = $zips | Select-Object -Skip $RetentionCount
        foreach ($z in $zipsToRemove) { Remove-Item $z.FullName -Force }
    }
    Write-Host "Backup gemaakt: $zipPath" -ForegroundColor Green
}

# Export backup
if ($config.backup.enableExportBackup -eq $true) {
    $ExportSource = "$ExportDir/*"
    New-ZipBackup -SourcePath $ExportSource -BackupFolder $BackupExportDir -BackupPrefix "export" -RetentionCount $config.backup.exportBackupRetention
}

# Archive backup
if ($config.backup.enableArchiveBackup -eq $true) {
    $ArchiveSource = "$ArchiveDir/*"
    New-ZipBackup -SourcePath $ArchiveSource -BackupFolder $BackupArchiveDir -BackupPrefix "archive" -RetentionCount $config.backup.archiveBackupRetention
}

# Config/Credentials backup
if ($config.backup.enableConfigBackup -eq $true) {
    $ConfigFiles = @()
    if (Test-Path -Path "./config.json" -PathType Leaf) { $ConfigFiles += "./config.json" }
    if (Test-Path -Path "./credentials.json" -PathType Leaf) { $ConfigFiles += "./credentials.json" }
    if ($ConfigFiles.Count -gt 0) {
        $ConfigZipName = "configcreds-$(Get-Date -Format 'yyyyMMdd').zip"
        $ConfigZipPath = Join-Path $BackupConfigDir $ConfigZipName
        Compress-Archive -Path $ConfigFiles -DestinationPath $ConfigZipPath -Force
        $zips = Get-ChildItem -Path $BackupConfigDir -Filter "configcreds-*.zip" | Sort-Object LastWriteTime -Descending
        if ($zips.Count -gt $config.backup.configBackupRetention) {
            $zipsToRemove = $zips | Select-Object -Skip $config.backup.configBackupRetention
            foreach ($z in $zipsToRemove) { Remove-Item $z.FullName -Force }
        }
        Write-Host "Backup gemaakt: $ConfigZipPath" -ForegroundColor Green
    }
}

Write-Host "Backups voltooid en opgeslagen in de backup directory." -ForegroundColor Cyan