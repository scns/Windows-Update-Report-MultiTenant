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
2025-10-16

.VERSIE
3.1.2
#>

#Versie-informatie
    $ProjectVersion = "3.1.2"
    $LastEditDate = "2025-10-16"

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

# Timezone conversie functie
function Convert-UTCToLocalTime {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UTCTimeString,
        [Parameter(Mandatory=$false)]
        [int]$OffsetHours = 0
    )
    
    try {
        # Controleer of de string al timezone info bevat
        if ($UTCTimeString -match 'Z$' -or $UTCTimeString -match '[+-]\d{2}:\d{2}$') {
            # Parse als UTC tijd met timezone info
            $utcTime = [DateTime]::Parse($UTCTimeString).ToUniversalTime()
        } else {
            # Probeer verschillende DateTime formaten
            $utcTime = $null
            $formats = @(
                "yyyy-MM-ddTHH:mm:ss.fffZ",
                "yyyy-MM-ddTHH:mm:ssZ", 
                "yyyy-MM-ddTHH:mm:ss",
                "MM/dd/yyyy HH:mm:ss",
                "dd/MM/yyyy HH:mm:ss"
            )
            
            foreach ($format in $formats) {
                try {
                    $utcTime = [DateTime]::ParseExact($UTCTimeString, $format, $null)
                    break
                } catch {
                    # Probeer volgende formaat
                }
            }
            
            # Als geen formaat werkt, probeer standaard parse
            if (-not $utcTime) {
                $utcTime = [DateTime]::Parse($UTCTimeString)
            }
        }
        
        # Voeg offset toe
        $localTime = $utcTime.AddHours($OffsetHours)
        
        # Return formatted string
        return $localTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        # Return original string als conversie faalt
        Write-Verbose "Timezone conversie gefaald voor '$UTCTimeString': $($_.Exception.Message)"
        return $UTCTimeString
    }
}

# Globale cache voor KB mapping
$Global:CachedKBMapping = $null
$Global:KBMappingCacheTime = $null
$Global:KBMappingCacheValidMinutes = 30  # Cache geldig voor 30 minuten

# Globale cache voor Office version mapping
$Global:CachedOfficeMapping = $null
$Global:OfficeMappingCacheTime = $null
$Global:OfficeMappingCacheValidMinutes = 30  # Cache geldig voor 30 minuten

# Timezone offset uit config (standaard 0 voor UTC)
$TimezoneOffsetHours = if ($config.timezoneOffsetHours) { $config.timezoneOffsetHours } else { 0 }
Write-Host "Timezone offset: UTC+$TimezoneOffsetHours uur" -ForegroundColor Cyan

# Functie om KB mapping te laden en cachen
function Get-CachedKBMapping {
    param(
        [string]$OnlineKBUrl,
        [int]$TimeoutSeconds = 10,
        [int]$CacheValidMinutes = 30
    )
    
    # Controleer of cache nog geldig is
    $now = Get-Date
    if ($Global:CachedKBMapping -and $Global:KBMappingCacheTime) {
        $cacheAge = ($now - $Global:KBMappingCacheTime).TotalMinutes
        if ($cacheAge -lt $CacheValidMinutes) {
            Write-Verbose "Using cached KB mapping (cached $([Math]::Round($cacheAge, 1)) minutes ago)"
            return @{
                Success = $true
                Data = $Global:CachedKBMapping
                Source = "Cache"
            }
        }
    }
    
    # Cache is verlopen of niet aanwezig, probeer online op te halen
    try {
        Write-Verbose "Fetching fresh KB mapping from: $OnlineKBUrl (timeout: $TimeoutSeconds seconds)"
        $onlineMapping = Invoke-RestMethod -Uri $OnlineKBUrl -Method GET -TimeoutSec $TimeoutSeconds -ErrorAction Stop
        
        # Update cache
        $Global:CachedKBMapping = $onlineMapping
        $Global:KBMappingCacheTime = $now
        
        Write-Verbose "KB mapping successfully cached (valid for $CacheValidMinutes minutes)"
        return @{
            Success = $true
            Data = $onlineMapping
            Source = "Online"
        }
    } catch {
        Write-Verbose "Failed to fetch online KB mapping: $($_.Exception.Message)"
        
        # Als er een oude cache is, gebruik die als fallback
        if ($Global:CachedKBMapping) {
            $cacheAge = ($now - $Global:KBMappingCacheTime).TotalMinutes
            Write-Verbose "Using expired cached KB mapping (cached $([Math]::Round($cacheAge, 1)) minutes ago)"
            return @{
                Success = $true
                Data = $Global:CachedKBMapping
                Source = "ExpiredCache"
            }
        }
        
        return @{
            Success = $false
            Data = $null
            Source = "Failed"
            Error = $_.Exception.Message
        }
    }
}

# Functie om Office version mapping te laden en cachen
function Get-CachedOfficeMapping {
    param(
        [string]$OnlineOfficeUrl,
        [string]$LocalOfficePath = ".\office-version-mapping.json",
        [int]$TimeoutSeconds = 10,
        [int]$CacheValidMinutes = 30
    )
    
    # Controleer of cache nog geldig is
    $now = Get-Date
    if ($Global:CachedOfficeMapping -and $Global:OfficeMappingCacheTime) {
        $cacheAge = ($now - $Global:OfficeMappingCacheTime).TotalMinutes
        if ($cacheAge -lt $CacheValidMinutes) {
            Write-Verbose "Using cached Office mapping (cached $([Math]::Round($cacheAge, 1)) minutes ago)"
            return @{
                Success = $true
                Data = $Global:CachedOfficeMapping
                Source = "Cache"
            }
        }
    }
    
    # Probeer eerst online op te halen
    if ($OnlineOfficeUrl) {
        try {
            Write-Verbose "Fetching fresh Office mapping from: $OnlineOfficeUrl (timeout: $TimeoutSeconds seconds)"
            $onlineMapping = Invoke-RestMethod -Uri $OnlineOfficeUrl -Method GET -TimeoutSec $TimeoutSeconds -ErrorAction Stop
            
            # Update cache
            $Global:CachedOfficeMapping = $onlineMapping
            $Global:OfficeMappingCacheTime = $now
            
            Write-Verbose "Office mapping successfully cached (valid for $CacheValidMinutes minutes)"
            return @{
                Success = $true
                Data = $onlineMapping
                Source = "Online"
            }
        } catch {
            Write-Verbose "Failed to fetch online Office mapping: $($_.Exception.Message)"
        }
    }
    
    # Probeer lokaal bestand te laden als fallback
    if (Test-Path $LocalOfficePath) {
        try {
            Write-Verbose "Loading Office mapping from local file: $LocalOfficePath"
            $localMapping = Get-Content -Path $LocalOfficePath -Raw | ConvertFrom-Json
            
            # Update cache
            $Global:CachedOfficeMapping = $localMapping
            $Global:OfficeMappingCacheTime = $now
            
            Write-Verbose "Office mapping successfully cached from local file (valid for $CacheValidMinutes minutes)"
            return @{
                Success = $true
                Data = $localMapping
                Source = "LocalFile"
            }
        } catch {
            Write-Verbose "Failed to load local Office mapping: $($_.Exception.Message)"
        }
    }
    
    # Als er een oude cache is, gebruik die als fallback
    if ($Global:CachedOfficeMapping) {
        $cacheAge = ($now - $Global:OfficeMappingCacheTime).TotalMinutes
        Write-Verbose "Using expired cached Office mapping (cached $([Math]::Round($cacheAge, 1)) minutes ago)"
        return @{
            Success = $true
            Data = $Global:CachedOfficeMapping
            Source = "ExpiredCache"
        }
    }
    
    return @{
        Success = $false
        Data = $null
        Source = "Failed"
        Error = "Could not load Office mapping from any source"
    }
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
            "https://raw.githubusercontent.com/scns/Windows-Update-Report-MultiTenant/refs/heads/main/kb-mapping.json" 
        }
        
        $KBNumber = $null
        $UpdateTitle = $null
        
        # Methode 1: Probeer online KB mapping op te halen (met cache)
        $OnlineMapping = $null
        $timeout = if ($Config.kbMapping.onlineTimeout) { $Config.kbMapping.onlineTimeout } else { 10 }
        $cacheValidMinutes = if ($Config.kbMapping.cacheValidMinutes) { $Config.kbMapping.cacheValidMinutes } else { 30 }
        
        $kbMappingResult = Get-CachedKBMapping -OnlineKBUrl $OnlineKBUrl -TimeoutSeconds $timeout -CacheValidMinutes $cacheValidMinutes
        if ($kbMappingResult.Success) {
            $OnlineMapping = $kbMappingResult.Data
            Write-Verbose "KB mapping loaded from: $($kbMappingResult.Source)"
            
            # Zoek in de juiste Windows versie sectie
            $MajorBuildNumber = [int]($TargetBuild -replace '\.\d+$', '')  # Verwijder minor build nummer
            $MappingSection = if ($MajorBuildNumber -lt 22000) {
                $OnlineMapping.mappings.windows10
            } elseif ($MajorBuildNumber -ge 26200) {
                $OnlineMapping.mappings.windows11_25h2
            } elseif ($MajorBuildNumber -ge 26000) {
                $OnlineMapping.mappings.windows11_24h2
            } else {
                $OnlineMapping.mappings.windows11_22h2
            }
            
            # Zoek exacte match (inclusief minor builds)
            $FoundMapping = $null
            $MajorBuild = $TargetBuild -replace '\.\d+$', ''  # Verwijder minor build nummer
            
            # Eerst: zoek exacte match voor volledige build (inclusief minor)
            if ($MappingSection.$TargetBuild) {
                $FoundMapping = $MappingSection.$TargetBuild
                Write-Verbose "Found exact build match for: $TargetBuild"
            }
            # Tweede: zoek in builds sub-object voor minor builds
            elseif ($MappingSection.$MajorBuild -and $MappingSection.$MajorBuild.builds -and $MappingSection.$MajorBuild.builds.$TargetBuild) {
                $FoundMapping = $MappingSection.$MajorBuild.builds.$TargetBuild
                Write-Verbose "Found minor build match for: $TargetBuild in $MajorBuild.builds"
            }
            # Derde: zoek major build als fallback
            elseif ($MappingSection.$MajorBuild) {
                $FoundMapping = $MappingSection.$MajorBuild
                Write-Verbose "Found major build match for: $MajorBuild (fallback from $TargetBuild)"
            }
            
            if ($FoundMapping) {
                $KBNumber = $FoundMapping.kb
                # Check of de online title al "for Windows" bevat
                $baseTitle = $FoundMapping.title
                if (-not $baseTitle) { $baseTitle = "Cumulative Update" }
                if ($baseTitle -notmatch "for Windows") {
                    $UpdateTitle = "$($FoundMapping.date) $baseTitle for $WindowsProduct ($KBNumber)"
                } else {
                    $UpdateTitle = "$($FoundMapping.date) $baseTitle ($KBNumber)"
                }
                
                # Voeg beschrijving toe als beschikbaar
                if ($FoundMapping.description) {
                    $UpdateTitle += " - $($FoundMapping.description)"
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
        } else {
            Write-Verbose "Failed to load KB mapping: $($kbMappingResult.Error)"
        }
        
        # Methode 2: Fallback naar lokale KB mapping
        if (-not $KBNumber -and ($Config.kbMapping.fallbackToLocalMapping -ne $false)) {
            Write-Verbose "Using fallback local KB mapping"
            # Gebruik bekende patronen voor recente builds (bijgewerkt tot september 2025)
            $LocalKBMappings = @{
                # Windows 11 24H2 (2024-2025) - September updates
                "26100.5074" = @{ KB = "KB5065522"; Date = "2025-09"; Title = "Cumulative Update (Minor)" }
                "26100" = @{ KB = "KB5064081"; Date = "2025-08"; Title = "Cumulative Update" }
                "26010" = @{ KB = "KB5064081"; Date = "2025-08"; Title = "Cumulative Update" }
                
                # Windows 11 23H2 (2023-2025) - September updates
                "22631.4249" = @{ KB = "KB5065522"; Date = "2025-09"; Title = "Cumulative Update (Minor)" }
                "22631" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                
                # Windows 11 22H2 (2022-2025) - September updates
                "22621.4249" = @{ KB = "KB5065522"; Date = "2025-09"; Title = "Cumulative Update (Minor)" }
                "22621" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                
                # Windows 11 21H2 (2021-2025)
                "22000" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                
                # Windows 10 22H2 (2022-2025) - September updates (laatste voor EOL)
                "19045.5073" = @{ KB = "KB5065522"; Date = "2025-09"; Title = "Cumulative Update (Minor) - Last before EOL" }
                "19045" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                "19044.5073" = @{ KB = "KB5065522"; Date = "2025-09"; Title = "Cumulative Update (Minor) - Last before EOL" }
                "19044" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                "19043" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                "19042" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                "19041" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
                
                # Windows 10 oudere versies (bijgewerkt naar augustus 2025)
                "18363" = @{ KB = "KB5063878"; Date = "2025-08"; Title = "Cumulative Update" }
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
            
            # Zoek naar exacte match eerst (inclusief minor builds)
            $ClosestMapping = $null
            $MajorBuild = $TargetBuild -replace '\.\d+$', ''  # Verwijder minor build nummer
            
            if ($LocalKBMappings.ContainsKey($TargetBuild)) {
                $ClosestMapping = $LocalKBMappings[$TargetBuild]
                Write-Verbose "Found exact local mapping for: $TargetBuild"
            } elseif ($LocalKBMappings.ContainsKey($MajorBuild)) {
                $ClosestMapping = $LocalKBMappings[$MajorBuild]
                Write-Verbose "Found major build local mapping for: $MajorBuild (from $TargetBuild)"
            } else {
                # Zoek naar de dichtstbijzijnde mapping
                $SmallestDifference = [int]::MaxValue
                
                foreach ($buildKey in $LocalKBMappings.Keys) {
                    # Skip minor builds bij het zoeken naar dichtstbijzijnde
                    if ($buildKey -match '\.\d+$') { continue }
                    
                    $difference = [Math]::Abs([int]$buildKey - [int]$MajorBuild)
                    if ($difference -lt $SmallestDifference) {
                        $SmallestDifference = $difference
                        $ClosestMapping = $LocalKBMappings[$buildKey]
                    }
                }
                Write-Verbose "Found closest local mapping with difference: $SmallestDifference"
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
            Method = if ($OnlineMapping -and $kbMappingResult.Source) { $kbMappingResult.Source } elseif ($KBNumber) { "Local" } else { "Estimated" }
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

# Functie om alleen KB nummers te extraheren uit Windows Update displayName
function Get-CleanUpdateIdentifier {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UpdateDisplayName
    )
    
    if (-not $UpdateDisplayName) {
        return ""
    }
    
    # Extract alleen KB nummers - meest voorkomende patroon
    if ($UpdateDisplayName -match '(KB\d{7})') {
        return $matches[1]
    }
    
    # Voor updates zonder KB nummer - probeer datum patroon voor identificatie
    if ($UpdateDisplayName -match '(\d{4}-\d{2})') {
        return "$($matches[1]) Update"
    }
    
    # Voor Defender updates zonder specifiek KB
    if ($UpdateDisplayName -match 'Defender|Security Intelligence') {
        return "Defender Update"
    }
    
    # Voor .NET Framework updates
    if ($UpdateDisplayName -match '\.NET.*?Framework') {
        return ".NET Framework Update"
    }
    
    # Voor Microsoft Edge updates
    if ($UpdateDisplayName -match 'Microsoft Edge') {
        return "Edge Update"
    }
    
    # Voor Office updates
    if ($UpdateDisplayName -match 'Microsoft Office') {
        return "Office Update"
    }
    
    # Als geen herkenbaar patroon - return lege string
    return ""
}

# Functie om Windows versie te bepalen op basis van build nummer
function Get-WindowsVersionFromBuild {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OSVersion
    )
    
    # Extract build number from OS version (e.g., "10.0.26100.2454" -> "26100")
    if ($OSVersion -match '10\.0\.(\d+)\.') {
        $buildNumber = [int]$matches[1]
        
        # Bepaal Windows versie op basis van build nummer ranges
        if ($buildNumber -ge 26200) {
            return "Windows 11 25H2"
        } elseif ($buildNumber -ge 26000 -and $buildNumber -le 26199) {
            return "Windows 11 24H2" 
        } elseif ($buildNumber -ge 22000 -and $buildNumber -le 25999) {
            return "Windows 11 22H2"
        } elseif ($buildNumber -ge 19000 -and $buildNumber -le 21999) {
            return "Windows 10"
        }
    }
    
    # Fallback voor onbekende versies
    return "Windows (Onbekend)"
}

# Functie om missing updates te bepalen op basis van KB database
function Get-MissingUpdatesFromKBDatabase {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CurrentOSVersion,
        [object]$KBMappingCache
    )
    
    $missingKBs = @()
    
    try {
        # Parse OS versie
        if ($CurrentOSVersion -notmatch '10\.0\.(\d+)\.(\d+)') {
            Write-Verbose "Could not parse OS version: $CurrentOSVersion"
            return @()
        }
        
        $majorBuild = $matches[1]
        $minorBuild = [int]$matches[2]
        $currentBuildString = "$majorBuild.$minorBuild"
        
        # Bepaal Windows versie en subversie (STRIKT gescheiden per versie-familie)
        if ([int]$majorBuild -lt 22000) {
            $windowsVersion = "windows10"
        } elseif ([int]$majorBuild -ge 26200) {
            $windowsVersion = "windows11_25h2"  # 25H2: alleen 26200+ builds
        } elseif ([int]$majorBuild -ge 26000) {
            $windowsVersion = "windows11_24h2"  # 24H2: alleen 26000-26199 builds  
        } else {
            $windowsVersion = "windows11_22h2"  # 22H2/23H2: alleen 22000-25999 builds
        }
        
        # Toegang tot KB mapping
        if (-not $KBMappingCache -or -not $KBMappingCache.mappings -or -not $KBMappingCache.mappings.$windowsVersion) {
            Write-Verbose "KB mapping cache not available for $windowsVersion"
            return @()
        }
        
        $versionMappings = $KBMappingCache.mappings.$windowsVersion
        
        # Zoek de juiste major build mapping
        if (-not $versionMappings.$majorBuild) {
            Write-Verbose "No KB mapping found for build $majorBuild in $windowsVersion"
            return @()
        }
        
        $buildMapping = $versionMappings.$majorBuild
        
        # Check of er specifieke build mappings zijn
        if ($buildMapping.builds) {
            $allBuilds = @()
            
            # KRITIEK: Alleen vergelijken binnen DEZELFDE Windows versie (24H2, 25H2, etc.)
            # Verzamel alleen builds die tot dezelfde major build EN Windows versie behoren
            foreach ($buildKey in $buildMapping.builds.PSObject.Properties.Name) {
                if ($buildKey -match "$majorBuild\.(\d+)") {
                    $buildMajor = [int]$majorBuild
                    $buildMinor = [int]$matches[1]
                    
                    # EXTRA VEILIGHEID: Valideer dat build echt bij deze Windows versie hoort
                    $buildBelongsToThisVersion = $false
                    switch ($windowsVersion) {
                        "windows10" { $buildBelongsToThisVersion = ($buildMajor -lt 22000) }
                        "windows11_22h2" { $buildBelongsToThisVersion = ($buildMajor -ge 22000 -and $buildMajor -lt 26000) }
                        "windows11_24h2" { $buildBelongsToThisVersion = ($buildMajor -ge 26000 -and $buildMajor -lt 26200) }
                        "windows11_25h2" { $buildBelongsToThisVersion = ($buildMajor -ge 26200) }
                    }
                    
                    if ($buildBelongsToThisVersion) {
                        $allBuilds += @{
                            Build = $buildKey
                            MinorVersion = $buildMinor
                            KB = $buildMapping.builds.$buildKey.kb
                            Date = $buildMapping.builds.$buildKey.date
                        }
                    } else {
                        Write-Verbose "Skipping build $buildKey - belongs to different Windows version"
                    }
                }
            }
            
            # Sorteer builds op minor versie (oplopend)
            $sortedBuilds = $allBuilds | Sort-Object MinorVersion
            
            # AANGEPAST: Vind alleen nieuwere builds binnen DEZELFDE major build reeks
            # Dit voorkomt dat 24H2 machines 25H2 updates als "missing" krijgen
            $targetBuilds = $sortedBuilds | Where-Object { $_.MinorVersion -gt $minorBuild }
            
            # Verzamel unieke KB nummers van alle hogere builds (binnen dezelfde OS versie)
            $uniqueKBs = @()
            foreach ($build in $targetBuilds) {
                if ($build.KB -and $uniqueKBs -notcontains $build.KB) {
                    $uniqueKBs += $build.KB
                }
            }
            
            $missingKBs = $uniqueKBs
            
            if ($missingKBs.Count -gt 0) {
                Write-Verbose "$windowsVersion - Current: $currentBuildString → Missing KBs: $($missingKBs -join ', ')"
            } else {
                Write-Verbose "$windowsVersion - Current: $currentBuildString → Up-to-date"
            }
        } else {
            # Geen specifieke builds beschikbaar - machine is waarschijnlijk up-to-date
            Write-Verbose "$windowsVersion - No specific builds available for $majorBuild, assuming up-to-date"
        }
        
    } catch {
        Write-Verbose "Error in Get-MissingUpdatesFromKBDatabase: $($_.Exception.Message)"
    }
    
    return $missingKBs
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

# Verzamel Office Version Mapping informatie voor HTML rapport
Write-Host "Ophalen Office Version Mapping informatie voor HTML rapport..." -ForegroundColor White
$OfficeMappingForHTML = $null
try {
    $onlineOfficeUrl = if ($config.officeMapping -and $config.officeMapping.officeMappingUrl) { 
        $config.officeMapping.officeMappingUrl 
    } else { 
        "https://raw.githubusercontent.com/scns/Windows-Update-Report-MultiTenant/refs/heads/main/office-version-mapping.json" 
    }
    
    $officeMappingResult = Get-CachedOfficeMapping -OnlineOfficeUrl $onlineOfficeUrl -LocalOfficePath ".\office-version-mapping.json" -TimeoutSeconds 10 -CacheValidMinutes 30
    
    if ($officeMappingResult.Success) {
        $OfficeMappingForHTML = @{
            Success = $true
            Method = $officeMappingResult.Source
            Data = $officeMappingResult.Data
            LastUpdated = if ($officeMappingResult.Data.metadata.lastUpdated) { $officeMappingResult.Data.metadata.lastUpdated } else { "Onbekend" }
            Version = if ($officeMappingResult.Data.metadata.version) { $officeMappingResult.Data.metadata.version } else { "N/A" }
        }
        Write-Host "Office Mapping succesvol geladen via $($officeMappingResult.Source)" -ForegroundColor Green
    } else {
        $OfficeMappingForHTML = @{
            Success = $false
            Method = "Failed"
            Error = $officeMappingResult.Error
            LastUpdated = "N/A"
            Version = "N/A"
        }
        Write-Host "Office Mapping kon niet worden geladen: $($officeMappingResult.Error)" -ForegroundColor Yellow
    }
} catch {
    $OfficeMappingForHTML = @{
        Success = $false
        Method = "Exception"
        Error = $_.Exception.Message
        LastUpdated = "N/A"
        Version = "N/A"
    }
    Write-Host "Fout bij laden Office Mapping: $($_.Exception.Message)" -ForegroundColor Red
}

# Verzamel KB Mapping informatie voor HTML rapport
Write-Host "Ophalen KB Mapping informatie voor HTML rapport..." -ForegroundColor White
$KBMappingForHTML = $null
try {
    $kbMappingResult = Get-CachedKBMapping -OnlineKBUrl $config.kbMapping.kbMappingUrl -TimeoutSeconds $config.kbMapping.timeoutSeconds -CacheValidMinutes $config.kbMapping.cacheValidMinutes
    if ($kbMappingResult.Success) {
        $totalEntries = 0
        if ($kbMappingResult.Data -and $kbMappingResult.Data.mappings) {
            # Tel alle Windows versie entries (inclusief minor builds)
            $windowsVersions = @("windows10", "windows11_22h2", "windows11_24h2", "windows11_25h2")
            
            foreach ($osVersion in $windowsVersions) {
                if ($kbMappingResult.Data.mappings.$osVersion) {
                    $totalEntries += ($kbMappingResult.Data.mappings.$osVersion.PSObject.Properties | Measure-Object).Count
                    # Tel minor builds
                    foreach ($build in $kbMappingResult.Data.mappings.$osVersion.PSObject.Properties.Name) {
                        $buildInfo = $kbMappingResult.Data.mappings.$osVersion.$build
                        if ($buildInfo.builds) {
                            $totalEntries += ($buildInfo.builds.PSObject.Properties | Measure-Object).Count
                        }
                    }
                }
            }
            # Tel Historical entries
            if ($kbMappingResult.Data.mappings.historical) {
                foreach ($year in $kbMappingResult.Data.mappings.historical.PSObject.Properties.Name) {
                    $totalEntries += ($kbMappingResult.Data.mappings.historical.$year.PSObject.Properties | Measure-Object).Count
                }
            }
        }
        
        $KBMappingForHTML = @{
            Success = $true
            Method = $kbMappingResult.Source
            Data = $kbMappingResult.Data
            LastUpdated = if ($kbMappingResult.Data.lastUpdated) { 
                "$($kbMappingResult.Data.lastUpdated) (v$($kbMappingResult.Data.version))" 
            } elseif ($Global:KBMappingCacheTime) { 
                $Global:KBMappingCacheTime.ToString("dd-MM-yyyy HH:mm:ss") 
            } else { 
                "N/A" 
            }
            TotalEntries = $totalEntries
        }
        Write-Host "KB Mapping geladen: $($KBMappingForHTML.TotalEntries) items via $($KBMappingForHTML.Method)" -ForegroundColor Green
    } else {
        $KBMappingForHTML = @{
            Success = $false
            Method = "Error"
            Error = $kbMappingResult.Error
            LastUpdated = "N/A"
            TotalEntries = 0
        }
        Write-Warning "KB Mapping kon niet worden geladen: $($kbMappingResult.Error)"
    }
} catch {
    $KBMappingForHTML = @{
        Success = $false
        Method = "Exception"
        Error = $_.Exception.Message
        LastUpdated = "N/A"
        TotalEntries = 0
    }
    Write-Warning "Fout bij ophalen KB Mapping: $($_.Exception.Message)"
}

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
                    LastSeen = Convert-UTCToLocalTime -UTCTimeString $_.LastSeen -OffsetHours $TimezoneOffsetHours
                    LoggedOnUsers = if ($_.LoggedOnUsers -is [System.Array]) { $_.LoggedOnUsers -join ', ' } else { $_.LoggedOnUsers }
                    OSPlatform = $_.OSPlatform
                    OSVersion = $_.OSVersion
                    OfficeVersion = "Niet beschikbaar (fallback mode)"
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
                    $ComplianceStatus = "Onbekend"
                    
                    # Haal Office versie op
                    $OfficeVersion = "Niet gedetecteerd"
                    try {
                        $DetectedAppsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($Device.id)')/detectedApps"
                        $DetectedApps = Invoke-MgGraphRequest -Method GET -Uri $DetectedAppsUri -ErrorAction SilentlyContinue
                        
                        if ($DetectedApps.value) {
                            # Zoek naar Office apps
                            $OfficeApp = $DetectedApps.value | Where-Object { 
                                $_.displayName -match 'Microsoft 365|Office 365|Microsoft Office|Office Professional|Office Standard'
                            } | Select-Object -First 1
                            
                            if ($OfficeApp -and $OfficeApp.version) {
                                $OfficeVersion = $OfficeApp.version
                            }
                        }
                    } catch {
                        Write-Verbose "Kon Office versie niet ophalen voor $DeviceName`: $($_.Exception.Message)"
                    }
                    
                    # Controleer Windows Update compliance status
                    $ComplianceUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($Device.id)')/deviceCompliancePolicyStates"
                    try {
                        $ComplianceStates = Invoke-MgGraphRequest -Method GET -Uri $ComplianceUri -ErrorAction SilentlyContinue
                        
                        $HasUpdateIssues = $false
                        $HasNonCompliant = $false
                        if ($ComplianceStates.value) {
                            foreach ($ComplianceState in $ComplianceStates.value) {
                                # Check voor overall compliance
                                if ($ComplianceState.state -eq 'nonCompliant') {
                                    $HasNonCompliant = $true
                                }
                                
                                # Check voor update specifieke issues
                                if ($ComplianceState.state -eq 'nonCompliant' -and 
                                    ($ComplianceState.settingName -like '*update*' -or 
                                     $ComplianceState.displayName -like '*Update*' -or
                                     $ComplianceState.displayName -like '*Windows*')) {
                                    $MissingUpdates += "Non-compliant: $($ComplianceState.displayName)"
                                    $HasUpdateIssues = $true
                                }
                            }
                            # Bepaal compliance status
                            $ComplianceStatus = if ($HasNonCompliant) { "Non-Compliant" } else { "Compliant" }
                        } else {
                            $ComplianceStatus = "Geen data"
                        }
                        
                        if ($HasUpdateIssues) {
                            $UpdateStatus = "Compliance problemen"
                        }
                    } catch {
                        Write-Verbose "Compliance check failed for $DeviceName"
                        $ComplianceStatus = "Error"
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
                            try {
                                # Probeer verschillende DateTime formaten voor robuuste parsing
                                $syncTime = $null
                                $formats = @(
                                    "yyyy-MM-ddTHH:mm:ss.fffZ",
                                    "yyyy-MM-ddTHH:mm:ssZ", 
                                    "yyyy-MM-ddTHH:mm:ss",
                                    "MM/dd/yyyy HH:mm:ss",
                                    "dd/MM/yyyy HH:mm:ss"
                                )
                                
                                foreach ($format in $formats) {
                                    try {
                                        $syncTime = [DateTime]::ParseExact($Device.lastSyncDateTime, $format, $null)
                                        break
                                    } catch { }
                                }
                                
                                if (-not $syncTime) {
                                    $syncTime = [DateTime]::Parse($Device.lastSyncDateTime)
                                }
                                
                                # Vergelijk met huidige tijd (beide in UTC als mogelijk)
                                $now = if ($syncTime.Kind -eq 'Utc') { (Get-Date).ToUniversalTime() } else { Get-Date }
                                [Math]::Round((New-TimeSpan -Start $syncTime -End $now).TotalDays)
                            } catch {
                                Write-Verbose "Kan lastSyncDateTime niet parsen voor $($Device.deviceName): $($Device.lastSyncDateTime)"
                                999
                            }
                        } else { 999 }
                        
                        # Controleer OS versie - voeg deze informatie toe voor transparantie
                        $OSVersion = if ($Device.osVersion) { $Device.osVersion } else { "Onbekend" }
                        
                        if ($DaysSinceSync -le 3) {
                            $MissingUpdates = @()  # Laat leeg voor up-to-date machines
                            $UpdateStatus = "Up to date"
                            
                            # Voor machines die recent hebben gesynchroniseerd, controleer via KB database
                            if ($OSVersion -and $OSVersion -match '10\.0\.(\d+)\.(\d+)') {
                                $currentBuild = "$($matches[1]).$($matches[2])"
                                
                                # Gebruik KB database om missing updates te bepalen
                                $kbMappingCache = $Global:CachedKBMapping
                                $missingKBsFromDB = Get-MissingUpdatesFromKBDatabase -CurrentOSVersion $OSVersion -KBMappingCache $kbMappingCache
                                
                                if ($missingKBsFromDB.Count -gt 0) {
                                    # Machine heeft updates nodig volgens KB database
                                    $ActualMissingUpdates += $missingKBsFromDB
                                    $UpdateStatus = "Update beschikbaar"
                                    Write-Verbose "Build $currentBuild needs updates: $($missingKBsFromDB -join ', ')"
                                } else {
                                    # Geen missing updates gevonden in KB database
                                    Write-Verbose "Build $currentBuild is up-to-date according to KB database"
                                }
                            }
                        } elseif ($DaysSinceSync -le 7) {
                            $MissingUpdates = @()  # Laat leeg voor waarschijnlijk up-to-date machines
                            $UpdateStatus = "Waarschijnlijk up to date"
                        } else {
                            $MissingUpdates = @("Windows Update status: Niet recent gesynchroniseerd ($DaysSinceSync dagen) (OS: $OSVersion)")
                            $UpdateStatus = "Synchronisatie vereist"
                        }
                    }
                    
                    # Bepaal het aantal missing updates op basis van unieke KB's in de lijst
                    $UpdateCount = if ($UpdateStatus -in @("Up to date", "Waarschijnlijk up to date")) { 
                        0 
                    } elseif ($ActualMissingUpdates -and $ActualMissingUpdates.Count -gt 0) { 
                        # Tel unieke KB nummers
                        $uniqueKBs = if ($ActualMissingUpdates -is [System.Array]) { 
                            $ActualMissingUpdates | Select-Object -Unique
                        } else { 
                            @($ActualMissingUpdates)
                        }
                        $uniqueKBs.Count
                    } else { 
                        1  # Fallback voor onbekende status
                    }
                    
                    $ResultsArray += [PSCustomObject]@{
                        DeviceName = $DeviceName
                        MissingUpdates = $MissingUpdates
                        ActualMissingUpdates = $ActualMissingUpdates
                        Count = $UpdateCount
                        LastSeen = Convert-UTCToLocalTime -UTCTimeString $Device.lastSyncDateTime -OffsetHours $TimezoneOffsetHours
                        LoggedOnUsers = if ($Device.userPrincipalName) { $Device.userPrincipalName } else { "Geen gebruiker" }
                        OSPlatform = $Device.operatingSystem
                        OSVersion = $Device.osVersion
                        OfficeVersion = $OfficeVersion
                        UpdateStatus = $UpdateStatus
                        ComplianceStatus = $ComplianceStatus
                    }
                    
                } catch {
                    Write-Warning "Fout bij verwerken device $($Device.deviceName): $($_.Exception.Message)"
                    
                    $ResultsArray += [PSCustomObject]@{
                        DeviceName = $Device.deviceName
                        MissingUpdates = @("Error: Kan Windows Update status niet controleren")
                        OfficeVersion = "Onbekend"
                        ActualMissingUpdates = @()
                        Count = 1  # Error = één probleem item
                        LastSeen = Convert-UTCToLocalTime -UTCTimeString $Device.lastSyncDateTime -OffsetHours $TimezoneOffsetHours
                        LoggedOnUsers = if ($Device.userPrincipalName) { $Device.userPrincipalName } else { "Geen gebruiker" }
                        OSPlatform = $Device.operatingSystem
                        OSVersion = if ($Device.osVersion) { $Device.osVersion } else { "Onbekend" }
                        UpdateStatus = "Error"
                        ComplianceStatus = "Error"
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
        # Gebruik versie-specifieke logica om cross-version vergelijking te voorkomen
        $OSVersionGroups = $ResultsArray | Where-Object { $_.OSVersion -and $_.OSVersion -ne "Onbekend" } | 
                          Group-Object OSVersion | 
                          Sort-Object Name -Descending
        
        # Bepaal nieuwste versie per Windows versie (22H2/24H2/25H2) om cross-version vergelijking te voorkomen
        $LatestVersionPerWindowsVersion = @{}
        
        foreach ($group in $OSVersionGroups) {
            $osVersion = $group.Name
            $windowsVersion = Get-WindowsVersionFromBuild $osVersion
            
            if (-not $LatestVersionPerWindowsVersion.ContainsKey($windowsVersion) -or 
                $osVersion -gt $LatestVersionPerWindowsVersion[$windowsVersion]) {
                $LatestVersionPerWindowsVersion[$windowsVersion] = $osVersion
            }
        }
        
        if ($OSVersionGroups.Count -gt 1) {
            $TotalMachines = ($OSVersionGroups | Measure-Object Count -Sum).Sum
            
            Write-Host "`n=== OS VERSIE ANALYSE ===" -ForegroundColor Yellow
            Write-Host "Totaal machines: $TotalMachines" -ForegroundColor Cyan
            
            # Toon nieuwste versie per Windows versie
            Write-Host "`nNieuwste versie per Windows versie:" -ForegroundColor Green
            foreach ($windowsVersion in $LatestVersionPerWindowsVersion.Keys | Sort-Object) {
                $latestBuild = $LatestVersionPerWindowsVersion[$windowsVersion]
                $machineCount = ($OSVersionGroups | Where-Object Name -eq $latestBuild).Count
                Write-Host "  $windowsVersion`: $latestBuild ($machineCount machines)" -ForegroundColor Green
            }
            
            Write-Host "`nVerdeling per OS versie:" -ForegroundColor Cyan
            
            foreach ($group in $OSVersionGroups) {
                $percentage = [Math]::Round(($group.Count / $TotalMachines) * 100, 1)
                $windowsVersion = Get-WindowsVersionFromBuild $group.Name
                $latestForThisVersion = $LatestVersionPerWindowsVersion[$windowsVersion]
                
                if ($group.Name -eq $latestForThisVersion) {
                    Write-Host "  [OK] $($group.Name): $($group.Count) machines ($percentage%) - $windowsVersion up-to-date" -ForegroundColor Green
                } else {
                    Write-Host "  [WARNING] $($group.Name): $($group.Count) machines ($percentage%) - $windowsVersion VEROUDERD" -ForegroundColor Yellow
                }
            }
            
            # Update de resultaten voor machines met verouderde OS versies - gebruik versie-specifieke logica
            $UpdatedResults = @()
            foreach ($result in $ResultsArray) {
                $newResult = $result.PSObject.Copy()
                $newResult | Add-Member -MemberType NoteProperty -Name "KBMethod" -Value "N/A" -Force
                
                # Bepaal de Windows versie van dit apparaat en de nieuwste versie voor die Windows versie
                if ($result.OSVersion -and $result.OSVersion -ne "Onbekend") {
                    $deviceWindowsVersion = Get-WindowsVersionFromBuild $result.OSVersion
                    $latestForThisWindowsVersion = $LatestVersionPerWindowsVersion[$deviceWindowsVersion]
                    
                    # Vergelijk alleen met de nieuwste versie binnen dezelfde Windows versie (22H2/24H2/25H2)
                    if ($result.OSVersion -ne $latestForThisWindowsVersion) {
                        # Machines met verouderde OS versie binnen hun Windows versie - update hun status
                        $originalStatus = $result.UpdateStatus
                        $newResult.UpdateStatus = "Verouderde OS versie"
                        
                        # Voeg informatie toe over de verouderde OS versie (binnen dezelfde Windows versie)
                        $versionDifference = "Huidige: $($result.OSVersion), Nieuwste voor $deviceWindowsVersion`: $latestForThisWindowsVersion"
                        $newResult.MissingUpdates += "LET OP: Verouderde OS versie - $versionDifference (was: $originalStatus)"
                        
                        # Probeer te bepalen welke cumulative update er nodig is op basis van OS versie
                        if ($result.OSVersion -and $latestForThisWindowsVersion) {
                            $currentBuild = $result.OSVersion -replace '.*\.(\d+)$', '$1'
                            $latestBuild = $latestForThisWindowsVersion -replace '.*\.(\d+)$', '$1'
                        
                            # Voor Windows 11/10 updates - haal KB nummers online op
                            if ($currentBuild -match '^\d+$' -and $latestBuild -match '^\d+$') {
                                $buildDifference = [int]$latestBuild - [int]$currentBuild
                                if ($buildDifference -gt 0) {
                                    # Gebruik KB database om missing updates te bepalen
                                    Write-Verbose "Looking up KB information for OS version: $($result.OSVersion)"
                                    $kbMappingCache = $Global:CachedKBMapping
                                    $missingKBsFromDB = Get-MissingUpdatesFromKBDatabase -CurrentOSVersion $result.OSVersion -KBMappingCache $kbMappingCache
                                    if ($missingKBsFromDB.Count -gt 0) {
                                        $newResult.ActualMissingUpdates += $missingKBsFromDB
                                        $newResult.KBMethod = "KB Database"
                                        Write-Verbose "Found missing KBs via KB Database: $($missingKBsFromDB -join ', ')"
                                    } else {
                                        # Geen updates gevonden in KB database
                                        $newResult.KBMethod = "KB Database (up-to-date)"
                                        Write-Verbose "No missing updates found in KB Database for $($result.OSVersion)"
                                    }
                                }
                            }
                        }
                        
                        # Bepaal het aantal missing updates voor verouderde OS versie (unieke KB's)
                        $OSUpdateCount = if ($newResult.ActualMissingUpdates -and $newResult.ActualMissingUpdates.Count -gt 0) { 
                            # Tel unieke KB nummers
                            $uniqueKBs = if ($newResult.ActualMissingUpdates -is [System.Array]) { 
                                $newResult.ActualMissingUpdates | Select-Object -Unique
                            } else { 
                                @($newResult.ActualMissingUpdates)
                            }
                            $uniqueKBs.Count
                        } else { 
                            1  # Minimaal één update voor verouderde OS versie
                        }
                        
                        $newResult.Count = $OSUpdateCount
                    }
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
                Count = 1  # Permission error = één probleem item
                LastSeen = (Get-Date).ToString()
                LoggedOnUsers = "N/A"
                OSPlatform = "Windows"
                OSVersion = "Onbekend" 
                UpdateStatus = "Permission Error"
                ComplianceStatus = "Error"
            })
        }
    }

    #Format the results into an array
    $ResultsTable = $Result.results | ForEach-Object {
        # Probeer eerst ActualMissingUpdates (echte KB nummers), anders MissingUpdates (status info)
        $MissingUpdatesDisplay = ""
        
        if ($_.ActualMissingUpdates -and $_.ActualMissingUpdates.Count -gt 0) {
            # Er zijn echte KB updates - verwijder duplicaten en toon deze
            $uniqueKBs = if ($_.ActualMissingUpdates -is [System.Array]) { 
                $_.ActualMissingUpdates | Select-Object -Unique
            } else { 
                @($_.ActualMissingUpdates)
            }
            $MissingUpdatesDisplay = $uniqueKBs -join '; '
        } elseif ($_.MissingUpdates -and $_.MissingUpdates.Count -gt 0) {
            # Geen KB updates, maar wel status informatie - toon eerste status
            $MissingUpdatesDisplay = if ($_.MissingUpdates -is [System.Array]) { 
                $_.MissingUpdates[0]  # Toon alleen de eerste status voor beknoptheid
            } else { 
                $_.MissingUpdates 
            }
        }
        
        [PSCustomObject]@{
            Device = $_.DeviceName
            "Update Status" = if ($_.UpdateStatus) { $_.UpdateStatus } else { "Onbekend" }
            "Compliance Status" = if ($_.ComplianceStatus) { $_.ComplianceStatus } else { "Onbekend" }
            "Missing Updates" = $MissingUpdatesDisplay
            "OS Version" = if ($_.OSVersion) { $_.OSVersion } else { "Onbekend" }
            "Office Version" = if ($_.OfficeVersion) { $_.OfficeVersion } else { "Niet gedetecteerd" }
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

# Verwijder trailing comma's en whitespace
$ChartDatasets = $ChartDatasets.TrimEnd(',', ' ', "`r", "`n", "`t")
$ChartDataJSON = $ChartDataJSON.TrimEnd(',') + "}"

# Bereken globale statistieken voor alle klanten
$GlobalStats = @{
    TotalPCs = 0
    UpToDatePCs = 0
    PendingPCs = 0
    FailedPCs = 0
    OutdatedPCs = 0
    ManualPCs = 0
    SyncPCs = 0
    CompliancePercentage = 0
}

foreach ($Customer in $LatestCsvPerCustomer.Keys) {
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
            $GlobalStats.TotalPCs++
            
            switch ($row.'Update Status') {
                { $_ -eq "Up to date" -or $_ -eq "Waarschijnlijk up to date" } {
                    $GlobalStats.UpToDatePCs++
                }
                { $_ -match "wachtend" } {
                    # PendingPCs wordt nu berekend via Count kolom (zie onder)
                }
                { $_ -eq "Synchronisatie vereist" } {
                    $GlobalStats.SyncPCs++
                }
                { $_ -match "fout|Error|problemen" } {
                    $GlobalStats.FailedPCs++
                }
                { $_ -eq "Verouderde OS versie" } {
                    $GlobalStats.OutdatedPCs++
                }
                { $_ -eq "Handmatige controle vereist" } {
                    $GlobalStats.ManualPCs++
                }
            }
            
            # PendingPCs is de som van alle Count waarden (1 = update vereist, 0 = up-to-date)
            $GlobalStats.PendingPCs += [int]$row.Count
        }
    }
}

# Bereken compliance percentage
if ($GlobalStats.TotalPCs -gt 0) {
    $GlobalStats.CompliancePercentage = [math]::Round(($GlobalStats.UpToDatePCs / $GlobalStats.TotalPCs) * 100)
}

# Genereer tabbladen en tabellen voor alleen de laatste datum per klant
$CustomerTabs = ""
$CustomerTables = ""
foreach ($Customer in ($LatestCsvPerCustomer.Keys | Sort-Object)) {
    $TableRows = ""
    $RowCount = 0
    
    # Bereken statistieken per klant
    $CustomerStats = @{
        TotalPCs = 0
        UpToDatePCs = 0
        PendingPCs = 0
        FailedPCs = 0
        OutdatedPCs = 0
        ManualPCs = 0
        SyncPCs = 0
        CompliancePercentage = 0
    }
    
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
            $CustomerStats.TotalPCs++
            
            # Tel statistieken per status
            switch ($row.'Update Status') {
                { $_ -eq "Up to date" -or $_ -eq "Waarschijnlijk up to date" } {
                    $CustomerStats.UpToDatePCs++
                }
                { $_ -match "wachtend" } {
                    # PendingPCs wordt nu berekend via Count kolom (zie onder)
                }
                { $_ -eq "Synchronisatie vereist" } {
                    $CustomerStats.SyncPCs++
                }
                { $_ -match "fout|Error|problemen" } {
                    $CustomerStats.FailedPCs++
                }
                { $_ -eq "Verouderde OS versie" } {
                    $CustomerStats.OutdatedPCs++
                }
                { $_ -eq "Handmatige controle vereist" } {
                    $CustomerStats.ManualPCs++
                }
            }
            
            # PendingPCs is de som van alle Count waarden (1 = update vereist, 0 = up-to-date)
            $CustomerStats.PendingPCs += [int]$row.Count
            
            # Voeg status kleuren toe
            $StatusColor = switch ($row.'Update Status') {
                "Up to date" { "color: #28a745; font-weight: bold;" }
                "Waarschijnlijk up to date" { "color: #28a745;" }
                "Verouderde OS versie" { "color: #fd7e14; font-weight: bold;" }
                "Compliance problemen" { "color: #dc3545; font-weight: bold;" }
                "Updates wachtend" { "color: #ffc107; font-weight: bold;" }
                "Update fouten" { "color: #dc3545; font-weight: bold;" }
                "Synchronisatie vereist" { "color: #17a2b8; font-weight: bold;" }
                "Error" { "color: #dc3545; font-weight: bold;" }
                "Handmatige controle vereist" { "color: #6f42c1;" }
                default { "color: inherit;" }
            }
            
            # Compliance Status kleuren
            $ComplianceColor = switch ($row.'Compliance Status') {
                "Compliant" { "color: #28a745;" }
                "Non-Compliant" { "color: #e74c3c; font-weight: bold;" }
                "Error" { "color: #fd7e14;" }
                "Geen data" { "color: #6c757d;" }
                "Onbekend" { "color: #6c757d;" }
                default { "color: #6c757d;" }
            }
            
            # Office Version kleuren bepalen op basis van mapping
            $OfficeColor = "color: #6c757d;"  # Default grijs voor niet gedetecteerd
            $OfficeVersionText = $row.'Office Version'
            $OfficeChannel = "Onbekend"
            
            if ($OfficeMappingForHTML.Success -and $OfficeVersionText -and $OfficeVersionText -ne "Niet gedetecteerd" -and $OfficeVersionText -ne "Niet beschikbaar (fallback mode)" -and $OfficeVersionText -ne "Onbekend") {
                try {
                    # Haal build nummer uit versie string (bijv. "16.0.19426.20186" → "19426")
                    if ($OfficeVersionText -match '16\.0\.(\d+)\.') {
                        $detectedBuild = $matches[1]
                        
                        # Vergelijk met Current Channel (nieuwste)
                        $currentChannelBuild = $null
                        if ($OfficeMappingForHTML.Data.update_channels.'Current Channel'.current_build) {
                            if ($OfficeMappingForHTML.Data.update_channels.'Current Channel'.current_build -match '^(\d+)\.') {
                                $currentChannelBuild = [int]$matches[1]
                            }
                        }
                        
                        # Vergelijk met Monthly Enterprise Channel (recent maar stabiel)
                        $monthlyEnterpriseBuild = $null
                        if ($OfficeMappingForHTML.Data.update_channels.'Monthly Enterprise Channel'.versions -and 
                            $OfficeMappingForHTML.Data.update_channels.'Monthly Enterprise Channel'.versions.Count -gt 0) {
                            $latestMonthly = $OfficeMappingForHTML.Data.update_channels.'Monthly Enterprise Channel'.versions[0]
                            if ($latestMonthly.build -match '^(\d+)\.') {
                                $monthlyEnterpriseBuild = [int]$matches[1]
                            }
                        }
                        
                        # Vergelijk met Semi-Annual (oudste ondersteunde)
                        $semiAnnualBuild = $null
                        if ($OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel'.versions -and 
                            $OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel'.versions.Count -gt 0) {
                            # Neem de oudste versie (laatste in de array)
                            $oldestSemiAnnual = $OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel'.versions[-1]
                            if ($oldestSemiAnnual.build -match '^(\d+)\.') {
                                $semiAnnualBuild = [int]$matches[1]
                            }
                        }
                        
                        $detectedBuildInt = [int]$detectedBuild
                        
                        # Bepaal kleur en channel op basis van build age
                        if ($currentChannelBuild -and $detectedBuildInt -ge $currentChannelBuild) {
                            # Nieuwste versie (Current Channel of nieuwer)
                            $OfficeColor = "color: #28a745; font-weight: bold;"  # Groen
                            $OfficeChannel = "Current Channel"
                        } elseif ($monthlyEnterpriseBuild -and $detectedBuildInt -ge $monthlyEnterpriseBuild) {
                            # Recent maar niet nieuwste (Monthly Enterprise)
                            $OfficeColor = "color: #28a745;"  # Groen maar niet bold
                            $OfficeChannel = "Monthly Enterprise"
                        } elseif ($semiAnnualBuild -and $detectedBuildInt -ge $semiAnnualBuild) {
                            # Oudere maar nog ondersteunde versie (Semi-Annual)
                            $OfficeColor = "color: #ffc107;"  # Oranje (waarschuwing)
                            $OfficeChannel = "Semi-Annual Enterprise"
                        } else {
                            # Zeer oude versie (mogelijk EOL)
                            $OfficeColor = "color: #dc3545; font-weight: bold;"  # Rood
                            $OfficeChannel = "Verouderd/EOL"
                        }
                    }
                } catch {
                    # Bij parse fouten, gebruik default grijs
                    $OfficeColor = "color: #6c757d;"
                    $OfficeChannel = "Onbekend"
                }
            }
            
            # Windows versie naam en support status bepalen
            $WindowsVersionName = ""
            $WindowsColor = "color: #6c757d;"  # Default grijs
            $OSVersionText = $row.'OS Version'
            
            if ($OSVersionText -and $OSVersionText -ne "Onbekend") {
                # Parse build number (bijv. "10.0.26100.2454" → "26100")
                if ($OSVersionText -match '10\.0\.(\d+)\.') {
                    $buildNumber = [int]$matches[1]
                    
                    # Windows 11 24H2 (26100)
                    if ($buildNumber -ge 26100 -and $buildNumber -lt 26200) {
                        $WindowsVersionName = "W11 24H2"
                        $WindowsColor = "color: #28a745; font-weight: bold;"  # Groen - In support tot oktober 2026
                    }
                    # Windows 11 25H2 (26200)
                    elseif ($buildNumber -ge 26200) {
                        $WindowsVersionName = "W11 25H2"
                        $WindowsColor = "color: #28a745; font-weight: bold;"  # Groen - Nieuwste versie
                    }
                    # Windows 11 23H2 (22631)
                    elseif ($buildNumber -ge 22631 -and $buildNumber -lt 26100) {
                        $WindowsVersionName = "W11 23H2"
                        $WindowsColor = "color: #28a745;"  # Groen - In support tot november 2025
                    }
                    # Windows 11 22H2 (22621)
                    elseif ($buildNumber -ge 22621 -and $buildNumber -lt 22631) {
                        $WindowsVersionName = "W11 22H2"
                        $WindowsColor = "color: #ffc107;"  # Oranje - EOL oktober 2024 (Pro), oktober 2025 (Enterprise)
                    }
                    # Windows 11 21H2 (22000)
                    elseif ($buildNumber -ge 22000 -and $buildNumber -lt 22621) {
                        $WindowsVersionName = "W11 21H2"
                        $WindowsColor = "color: #dc3545; font-weight: bold;"  # Rood - EOL oktober 2023
                    }
                    # Windows 10 22H2 (19045)
                    elseif ($buildNumber -ge 19045) {
                        $WindowsVersionName = "W10 22H2"
                        $WindowsColor = "color: #ffc107;"  # Oranje - EOL oktober 2025
                    }
                    # Windows 10 21H2 (19044)
                    elseif ($buildNumber -ge 19044 -and $buildNumber -lt 19045) {
                        $WindowsVersionName = "W10 21H2"
                        $WindowsColor = "color: #dc3545; font-weight: bold;"  # Rood - EOL juni 2023
                    }
                    # Windows 10 21H1 (19043)
                    elseif ($buildNumber -ge 19043 -and $buildNumber -lt 19044) {
                        $WindowsVersionName = "W10 21H1"
                        $WindowsColor = "color: #dc3545; font-weight: bold;"  # Rood - EOL december 2022
                    }
                    # Windows 10 20H2 (19042)
                    elseif ($buildNumber -ge 19042 -and $buildNumber -lt 19043) {
                        $WindowsVersionName = "W10 20H2"
                        $WindowsColor = "color: #dc3545; font-weight: bold;"  # Rood - EOL mei 2022
                    }
                    # Oudere Windows 10 versies
                    elseif ($buildNumber -ge 10240 -and $buildNumber -lt 19042) {
                        $WindowsVersionName = "W10 (oud)"
                        $WindowsColor = "color: #dc3545; font-weight: bold;"  # Rood - EOL
                    }
                    else {
                        $WindowsVersionName = "Onbekend"
                        $WindowsColor = "color: #6c757d;"
                    }
                }
            }
            
            $TableRows += "<tr><td>$($row.Device)</td><td style='$StatusColor'>$($row.'Update Status')</td><td style='$ComplianceColor'>$($row.'Compliance Status')</td><td>$($row.'Missing Updates')</td><td style='$WindowsColor'>$OSVersionText</td><td>$WindowsVersionName</td><td style='$OfficeColor'>$($row.'Office Version')</td><td>$OfficeChannel</td><td>$($row.Count)</td><td>$($row.LastSeen)</td><td>$($row.LoggedOnUsers)</td></tr>`n"
            $RowCount++
        }
    }
    
    # Bereken compliance percentage voor deze klant
    if ($CustomerStats.TotalPCs -gt 0) {
        $CustomerStats.CompliancePercentage = [math]::Round(($CustomerStats.UpToDatePCs / $CustomerStats.TotalPCs) * 100)
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
        
        <!-- Klant-specifieke statistieken -->
        <div class="global-stats">
            <div class="stat-card">
                <h4>Totaal PC's</h4>
                <div class="stat-number">$($CustomerStats.TotalPCs)</div>
            </div>
            <div class="stat-card up-to-date">
                <h4>Up to Date</h4>
                <div class="stat-number">$($CustomerStats.UpToDatePCs)</div>
            </div>
            <div class="stat-card pending">
                <h4>Updates Beschikbaar</h4>
                <div class="stat-number">$($CustomerStats.PendingPCs)</div>
            </div>
            <div class="stat-card failed">
                <h4>Update Fouten</h4>
                <div class="stat-number">$($CustomerStats.FailedPCs)</div>
            </div>
            <div class="stat-card info">
                <h4>Up to date %</h4>
                <div class="stat-number">$($CustomerStats.CompliancePercentage)%</div>
            </div>
        </div>
        
        <!-- Filter knoppen -->
        <div style="margin: 20px 0; background: #f8f9fa; padding: 15px; border-radius: 8px;">
            <h4 style="margin: 0 0 15px 0; color: #495057;"><i class="fa-solid fa-filter"></i> Filter Opties</h4>
            <div class="filter-buttons">
                <button class="filter-btn active" onclick="filterByStatus('overviewTable_$Customer', '')">
                    <i class="fa-solid fa-list"></i> Alle Statussen ($($CustomerStats.TotalPCs))
                </button>
                <button class="filter-btn up-to-date" onclick="filterByStatus('overviewTable_$Customer', 'Up to date')">
                    <i class="fa-solid fa-check-circle"></i> Up to date ($($CustomerStats.UpToDatePCs))
                </button>
                <button class="filter-btn failed" onclick="filterByStatus('overviewTable_$Customer', 'fout')">
                    <i class="fa-solid fa-exclamation-triangle"></i> Update Fouten ($($CustomerStats.FailedPCs))
                </button>
                <button class="filter-btn outdated" onclick="filterByStatus('overviewTable_$Customer', 'Verouderde OS versie')">
                    <i class="fa-solid fa-desktop"></i> Verouderde OS ($($CustomerStats.OutdatedPCs))
                </button>
                <button class="filter-btn manual" onclick="filterByStatus('overviewTable_$Customer', 'Handmatige controle vereist')">
                    <i class="fa-solid fa-user-cog"></i> Handmatige Controle ($($CustomerStats.ManualPCs))
                </button>
                <button class="filter-btn sync" onclick="filterByStatus('overviewTable_$Customer', 'Synchronisatie vereist')">
                    <i class="fa-solid fa-sync"></i> Synchronisatie Vereist ($($CustomerStats.SyncPCs))
                </button>
                <button class="filter-btn non-compliant" onclick="filterByCompliance('overviewTable_$Customer', 'Non-Compliant')">
                    <i class="fa-solid fa-times-circle"></i> Non-Compliant
                </button>
            </div>
        </div>
        
        <button onclick="exportTableToCSV('overviewTable_$Customer', '$Customer-full.csv', false)">Exporteren volledige tabel</button>
        <button onclick="exportTableToCSV('overviewTable_$Customer', '$Customer-filtered.csv', true)">Exporteren gefilterde rijen</button>
        <table id="overviewTable_$Customer" class="display" style="width:100%">
            <thead>
                <tr>
                    <th>Device</th>
                    <th>Update Status</th>
                    <th>Compliance Status</th>
                    <th>Missing Updates</th>
                    <th>OS Version</th>
                    <th>Windows Edition</th>
                    <th>Office Version</th>
                    <th>Office Channel</th>
                    <th>Count</th>
                    <th>LastSeen (UTC+$TimezoneOffsetHours)</th>
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
                "order": [[1, "desc"]],
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
                // Voor "Compliance Status" kolom (index 2): gebruik dropdown filter
                } else if (i === 2 && title.includes('Compliance Status')) {
                    // Verzamel unieke waarden uit de kolom
                    var uniqueValues = [];
                    table.column(i).data().unique().sort().each(function (d, j) {
                        if (d && uniqueValues.indexOf(d) === -1) {
                            uniqueValues.push(d);
                        }
                    });
                    
                    // Maak dropdown met unieke waarden
                    var select = '<br><select style="width:90%;font-size:12px;"><option value="">Alle compliance</option>';
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
        
        // Haal de klant naam uit de table ID (bijv. 'overviewTable_CustomerName')
        var customerName = tableId.replace('overviewTable_', '');
        
        // Update filter button states
        var filterContainer = document.querySelector('#' + customerName + ' .filter-buttons');
        if (filterContainer) {
            // Reset alle buttons
            filterContainer.querySelectorAll('.filter-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            
            // Activeer de juiste button
            var activeBtn = null;
            if (status === '') {
                activeBtn = filterContainer.querySelector('.filter-btn[onclick*="\'\'"]');
            } else {
                activeBtn = filterContainer.querySelector('.filter-btn[onclick*="' + status + '"]');
            }
            if (activeBtn) {
                activeBtn.classList.add('active');
            }
        }
        
        // Reset Compliance Status filter (kolom 2) wanneer Update Status filter wordt gebruikt
        table.column(2).search('').draw();
        $('#' + tableId + ' thead th:eq(2) select').val('');
        
        // Update Status kolom is index 1
        if (status === '') {
            // Reset filter - toon alle statussen
            table.column(1).search('').draw();
            // Reset ook de dropdown
            $('#' + tableId + ' thead th:eq(1) select').val('');
        } else {
            // Filter op specifieke status of gedeeltelijke match
            var searchTerm = status;
            if (status === 'wachtend') {
                searchTerm = 'wachtend|Synchronisatie';
            } else if (status === 'fout') {
                searchTerm = 'fout|Error|problemen';
            }
            table.column(1).search(searchTerm, true, false).draw();
            // Update ook de dropdown indien exacte match
            var exactMatch = table.column(1).data().toArray().find(d => d === status);
            if (exactMatch) {
                $('#' + tableId + ' thead th:eq(1) select').val(status);
            }
        }
    }
    
    // Functie voor snelfilters op Compliance Status
    window.filterByCompliance = function(tableId, complianceStatus) {
        var table = $('#' + tableId).DataTable();
        
        // Haal de klant naam uit de table ID
        var customerName = tableId.replace('overviewTable_', '');
        
        // Reset alle andere filter button states
        var updateFilterContainer = document.querySelector('#' + customerName + ' .filter-buttons');
        if (updateFilterContainer) {
            updateFilterContainer.querySelectorAll('.filter-btn').forEach(btn => {
                btn.classList.remove('active');
            });
        }
        
        // Reset Update Status filter (kolom 1)
        table.column(1).search('').draw();
        $('#' + tableId + ' thead th:eq(1) select').val('');
        
        // Compliance Status kolom is index 2
        if (complianceStatus === '') {
            // Reset filter - toon alle compliance statussen
            table.column(2).search('').draw();
            // Reset ook de dropdown
            $('#' + tableId + ' thead th:eq(2) select').val('');
        } else {
            // Filter op specifieke compliance status
            table.column(2).search('^' + complianceStatus + '$', true, false).draw();
            // Update ook de dropdown
            $('#' + tableId + ' thead th:eq(2) select').val(complianceStatus);
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
    
    /* Header styling  */
    .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; }
    .header h1 { margin: 0; color: #0066cc; }
    
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
    body.darkmode .header h1 { color: #66aaff; }
    body.darkmode canvas { background: #222; }
    body.darkmode table.dataTable thead th { background: #222; color: #eee; }
    body.darkmode .tab { border-bottom: 1px solid #444; }
    body.darkmode .tab button { background-color: #222; color: #eee; }
    body.darkmode .tab button:hover { background-color: #333; }
    body.darkmode .tab button.active { background-color: #444; }
    body.darkmode .tabcontent { background: #181a1b; color: #eee; }
    body.darkmode .footer { border-top: 1px solid #444; color: #aaa; }
    body.darkmode .footer a { color: #66aaff; }
    
    /* KB Database styling */
    .kb-status-box { margin-bottom: 20px; padding: 15px; background-color: #f5f5f5; border-left: 4px solid #0066cc; border-radius: 4px; }
    body.darkmode .kb-status-box { background-color: #2a2a2a; border-left-color: #66aaff; }
    
    .kb-warnings { background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 10px; margin: 10px 0; border-radius: 5px; }
    body.darkmode .kb-warnings { background-color: #4a3e1f; border-color: #6c5c00; color: #fff8dc; }
    body.darkmode .kb-warnings h3 { color: #ffc107; }
    
    .kb-preview { background-color: #d1ecf1; border: 1px solid #bee5eb; padding: 10px; margin: 10px 0; border-radius: 5px; }
    body.darkmode .kb-preview { background-color: #1f3a42; border-color: #356069; color: #e1f5fe; }
    body.darkmode .kb-preview h3 { color: #17a2b8; }
    
    /* KB Table hierarchical styling */
    .group-header { background-color: #e9ecef !important; font-weight: bold; }
    body.darkmode .group-header { background-color: #3a3a3a !important; color: #fff; }
    .group-header td { padding: 12px 8px !important; border-bottom: 2px solid #007bff !important; }
    body.darkmode .group-header td { border-bottom-color: #66aaff !important; }
    
    .build-row { background-color: #ffffff !important; }
    body.darkmode .build-row { background-color: #2d3035 !important; color: #ccc; }
    .build-row:hover { background-color: #f8f9fa !important; }
    body.darkmode .build-row:hover { background-color: #3d4045 !important; }
    
    /* Version badges */
    .version-badge { 
        padding: 4px 8px; 
        border-radius: 12px; 
        font-size: 11px; 
        font-weight: bold; 
        text-transform: uppercase; 
        letter-spacing: 0.5px;
    }
    .version-badge.win11-25h2 { background-color: #00d4aa; color: white; }
    .version-badge.win11-24h2 { background-color: #0078d4; color: white; }
    .version-badge.win11-22h2 { background-color: #5856d6; color: white; }
    .version-badge.win10 { background-color: #ff8c00; color: white; }
    .version-badge.historical { background-color: #6c757d; color: white; }
    body.darkmode .version-badge.win11-25h2 { background-color: #00a085; }
    body.darkmode .version-badge.win11-24h2 { background-color: #106ebe; }
    body.darkmode .version-badge.win11-22h2 { background-color: #4240b8; }
    body.darkmode .version-badge.win10 { background-color: #cc7000; }
    body.darkmode .version-badge.historical { background-color: #5a6268; }
    
    /* Snelfilter knoppen styling */
    .filter-container { margin-bottom: 15px; padding: 10px; background-color: #f8f9fa; border-radius: 5px; }
    body.darkmode .filter-container { background-color: #2d3035; }
    .filter-button { margin: 2px; padding: 5px 10px; border: none; border-radius: 3px; cursor: pointer; font-size: 12px; }
    .filter-button:hover { opacity: 0.8; }
    
    /* Filter Buttons  */
    .filter-buttons { display: flex; flex-wrap: wrap; gap: 10px; }
    .filter-btn { 
        background: white; 
        border: 2px solid #dee2e6; 
        padding: 8px 16px; 
        border-radius: 6px; 
        cursor: pointer; 
        font-size: 14px; 
        font-weight: 500;
        transition: all 0.3s;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .filter-btn:hover { 
        background: #f8f9fa; 
        border-color: #0066cc; 
        transform: translateY(-1px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .filter-btn.active { 
        background: #0066cc; 
        color: white; 
        border-color: #0066cc; 
    }
    .filter-btn.up-to-date { border-color: #28a745; color: #28a745; }
    .filter-btn.up-to-date:hover, .filter-btn.up-to-date.active { background: #28a745; color: white; }
    .filter-btn.pending { border-color: #ffc107; color: #e68900; }
    .filter-btn.pending:hover, .filter-btn.pending.active { background: #ffc107; color: #212529; }
    .filter-btn.failed { border-color: #dc3545; color: #dc3545; }
    .filter-btn.failed:hover, .filter-btn.failed.active { background: #dc3545; color: white; }
    .filter-btn.outdated { border-color: #fd7e14; color: #fd7e14; }
    .filter-btn.outdated:hover, .filter-btn.outdated.active { background: #fd7e14; color: white; }
    .filter-btn.manual { border-color: #6f42c1; color: #6f42c1; }
    .filter-btn.manual:hover, .filter-btn.manual.active { background: #6f42c1; color: white; }
    .filter-btn.sync { border-color: #17a2b8; color: #17a2b8; }
    .filter-btn.sync:hover, .filter-btn.sync.active { background: #17a2b8; color: white; }
    .filter-btn.non-compliant { border-color: #e74c3c; color: #e74c3c; font-weight: bold; }
    .filter-btn.non-compliant:hover, .filter-btn.non-compliant.active { background: #e74c3c; color: white; font-weight: bold; }
    
    /* Dark mode filter buttons */
    body.darkmode .filter-btn { background: #1e1e1e; border-color: #333; color: #e0e0e0; }
    body.darkmode .filter-btn:hover { background: #333; border-color: #4da6ff; }
    body.darkmode .filter-btn.active { background: #0066cc; color: white; border-color: #0066cc; }
    body.darkmode .filter-btn.up-to-date { border-color: #4dff4d; color: #4dff4d; }
    body.darkmode .filter-btn.up-to-date:hover, body.darkmode .filter-btn.up-to-date.active { background: #4dff4d; color: #121212; }
    body.darkmode .filter-btn.pending { border-color: #ffcc4d; color: #ffcc4d; }
    body.darkmode .filter-btn.pending:hover, body.darkmode .filter-btn.pending.active { background: #ffcc4d; color: #121212; }
    body.darkmode .filter-btn.failed { border-color: #ff4d4d; color: #ff4d4d; }
    body.darkmode .filter-btn.failed:hover, body.darkmode .filter-btn.failed.active { background: #ff4d4d; color: #121212; }
    body.darkmode .filter-btn.outdated { border-color: #ffaa66; color: #ffaa66; }
    body.darkmode .filter-btn.outdated:hover, body.darkmode .filter-btn.outdated.active { background: #ffaa66; color: #121212; }
    body.darkmode .filter-btn.manual { border-color: #b084ff; color: #b084ff; }
    body.darkmode .filter-btn.manual:hover, body.darkmode .filter-btn.manual.active { background: #b084ff; color: #121212; }
    body.darkmode .filter-btn.sync { border-color: #4dd4ff; color: #4dd4ff; }
    body.darkmode .filter-btn.sync:hover, body.darkmode .filter-btn.sync.active { background: #4dd4ff; color: #121212; }
    body.darkmode .filter-btn.non-compliant { border-color: #ff6b6b; color: #ff6b6b; font-weight: bold; }
    body.darkmode .filter-btn.non-compliant:hover, body.darkmode .filter-btn.non-compliant.active { background: #ff6b6b; color: #121212; font-weight: bold; }
    
    /* Dark mode container styling */
    body.darkmode div[style*="background: #f8f9fa"] { background: #2a2a2a !important; }
    
    /* Button Styling */
    #darkModeToggle { 
        background: #0066cc; 
        color: white; 
        border: none; 
        padding: 10px 15px; 
        border-radius: 5px; 
        cursor: pointer; 
        font-size: 14px;
        transition: background 0.3s;
    }
    #darkModeToggle:hover { background: #0056b3; }
    body.darkmode #darkModeToggle { background: #ffc107; color: #000; }
    body.darkmode #darkModeToggle:hover { background: #e0a800; }
    
    /* Statistieken kaarten styling */
    .global-stats { display: flex; flex-wrap: wrap; gap: 15px; margin: 20px 0; }
    .stat-card { 
        background: white; 
        padding: 20px; 
        border-radius: 8px; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); 
        text-align: center; 
        border-left: 4px solid #0066cc;
        flex: 1 1 150px;
        min-width: 150px;
    }
    .stat-card h4 { 
        margin: 0 0 10px 0; 
        font-size: 14px; 
        color: #666; 
        text-transform: uppercase; 
        letter-spacing: 0.5px; 
    }
    .stat-card .stat-number { 
        font-size: 28px; 
        font-weight: bold; 
        color: #333; 
    }
    .stat-card.up-to-date { border-left-color: #28a745; }
    .stat-card.up-to-date .stat-number { color: #28a745; }
    .stat-card.pending { border-left-color: #ffc107; }
    .stat-card.pending .stat-number { color: #e68900; }
    .stat-card.failed { border-left-color: #dc3545; }
    .stat-card.failed .stat-number { color: #dc3545; }
    .stat-card.info { border-left-color: #17a2b8; }
    .stat-card.info .stat-number { color: #17a2b8; }
    
    /* Dark mode statistieken styling */
    body.darkmode .stat-card { 
        background: #1e1e1e; 
        color: #e0e0e0; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.3); 
    }
    body.darkmode .stat-card h4 { color: #b0b0b0; }
    body.darkmode .stat-card .stat-number { color: #e0e0e0; }
    body.darkmode .stat-card.up-to-date .stat-number { color: #4dff4d; }
    body.darkmode .stat-card.pending .stat-number { color: #ffcc4d; }
    body.darkmode .stat-card.failed .stat-number { color: #ff4d4d; }
    body.darkmode .stat-card.info .stat-number { color: #4dffff; }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1><i class="fa-solid fa-desktop"></i> Windows Update Overview</h1>
        <button id="darkModeToggle"><i class="fa-solid fa-moon"></i> Dark mode</button>
    </div>

    <p><i class="fa-solid fa-clock"></i> Laatst uitgevoerd op: $LastRunDate</p>
    
    <!-- Statistieken sectie -->
    <div class="global-stats">
        <div class="stat-card">
            <h4>Totaal PC's</h4>
            <div class="stat-number">$($GlobalStats.TotalPCs)</div>
        </div>
        <div class="stat-card up-to-date">
            <h4>Up to Date</h4>
            <div class="stat-number">$($GlobalStats.UpToDatePCs)</div>
        </div>
        <div class="stat-card pending">
            <h4>Updates Beschikbaar</h4>
            <div class="stat-number">$($GlobalStats.PendingPCs)</div>
        </div>
        <div class="stat-card failed">
            <h4>Update Fouten</h4>
            <div class="stat-number">$($GlobalStats.FailedPCs)</div>
        </div>
        <div class="stat-card info">
            <h4>Up to date %</h4>
            <div class="stat-number">$($GlobalStats.CompliancePercentage)%</div>
        </div>
    </div>
    
    <h2>Totale Count per dag per klant</h2>
    <canvas id="countChart" height="100"></canvas>
    <div class="tab">
        <button class="tablinks" onclick="showAllCustomers()">Alle klanten</button>
        <button class="tablinks" onclick="showAppRegistrations()">App Registrations</button>
        <button class="tablinks" onclick="showKBMapping()">KB Mapping Database</button>
        <button class="tablinks" onclick="showOfficeVersions()">Office Versions</button>
        $CustomerTabs
    </div>
    
    <!-- App Registrations Tab -->
    <div id="AppRegistrations" class="tabcontent" style="display:none">
        <h2>App Registration Status Overview</h2>
        <div class="export-buttons" style="margin-bottom: 10px;">
            <button onclick="exportTableToCSV('appRegTable', 'App_Registrations-full.csv', false)">Exporteren volledige tabel</button>
            <button onclick="exportTableToCSV('appRegTable', 'App_Registrations-filtered.csv', true)">Exporteren gefilterde rijen</button>
        </div>
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
    
    <!-- KB Mapping Database Tab -->
    <div id="KBMapping" class="tabcontent" style="display:none">
        <h2>KB Mapping Database Overview</h2>
        <div class="kb-status-box">
            <strong>Database Status:</strong> $(if ($KBMappingForHTML.Success) { "<span style='color:green;'>OK Beschikbaar</span>" } else { "<span style='color:red;'>X Niet beschikbaar</span>" })<br>
            <strong>Bron Methode:</strong> $($KBMappingForHTML.Method)<br>
            <strong>Totaal Entries:</strong> $($KBMappingForHTML.TotalEntries)<br>
            <strong>Laatste Update:</strong> $($KBMappingForHTML.LastUpdated)
            $(if (-not $KBMappingForHTML.Success) { "<br><strong>Fout:</strong> <span style='color:red;'>$($KBMappingForHTML.Error)</span>" })
        </div>
        
        $(if ($KBMappingForHTML.Success -and $KBMappingForHTML.Data) {
            $warningsHtml = ""
            $previewHtml = ""
            
            # Toon warnings
            if ($KBMappingForHTML.Data.warnings) {
                $warningsHtml = "<div class='kb-warnings'>
                    <h3>WARNING Belangrijke Meldingen</h3>"
                foreach ($warning in $KBMappingForHTML.Data.warnings.PSObject.Properties) {
                    $warningsHtml += "<div style='margin:5px 0;'><strong>$($warning.Name):</strong> $($warning.Value)</div>"
                }
                $warningsHtml += "</div>"
            }
            
            # Toon preview informatie
            if ($KBMappingForHTML.Data.preview) {
                $previewHtml = "<div class='kb-preview'>
                    <h3>PREVIEW Aankomende Updates</h3>"
                foreach ($previewPeriod in $KBMappingForHTML.Data.preview.PSObject.Properties) {
                    $previewData = $previewPeriod.Value
                    $previewHtml += "<div style='margin:10px 0;'><strong>$($previewPeriod.Name.Replace('_', ' ')):</strong><br>"
                    if ($previewData.release_date) {
                        $previewHtml += "Release: $($previewData.release_date)<br>"
                    }
                    if ($previewData.type) {
                        $previewHtml += "Type: $($previewData.type)<br>"
                    }
                    foreach ($item in $previewData.PSObject.Properties) {
                        if ($item.Name -notin @('release_date', 'type')) {
                            if ($item.Value -is [PSCustomObject]) {
                                $previewHtml += "- $($item.Name): Build $($item.Value.build) &rArr; $($item.Value.kb)"
                                if ($item.Value.note) { $previewHtml += " ($($item.Value.note))" }
                                $previewHtml += "<br>"
                            } else {
                                $previewHtml += "- $($item.Name): $($item.Value)<br>"
                            }
                        }
                    }
                    $previewHtml += "</div>"
                }
                $previewHtml += "</div>"
            }
            
            "$warningsHtml$previewHtml"
        })
        
        $(if ($KBMappingForHTML.Success -and $KBMappingForHTML.Data) {
            $kbEntries = @()
            
            # Export buttons HTML
            $exportButtonsHtml = "<div class='export-buttons' style='margin-bottom: 10px;'>
                <button onclick=`"exportTableToCSV('kbMappingTable', 'KB_Mapping_Database-full.csv', false)`">Exporteren volledige tabel</button>
                <button onclick=`"exportTableToCSV('kbMappingTable', 'KB_Mapping_Database-filtered.csv', true)`">Exporteren gefilterde rijen</button>
            </div>"
            
            # Verwerk Windows 11 25H2 mappings
            if ($KBMappingForHTML.Data.mappings.windows11_25h2) {
                foreach ($baseBuild in ($KBMappingForHTML.Data.mappings.windows11_25h2.PSObject.Properties.Name | Sort-Object -Descending)) {
                    $kbInfo = $KBMappingForHTML.Data.mappings.windows11_25h2.$baseBuild
                    
                    # Specifieke builds (indien beschikbaar)
                    if ($kbInfo.builds) {
                        foreach ($specificBuild in ($kbInfo.builds.PSObject.Properties.Name | Sort-Object -Descending)) {
                            $buildInfo = $kbInfo.builds.$specificBuild
                            $kbEntries += "<tr class='build-row'><td>$specificBuild</td><td>$($buildInfo.kb)</td><td>$($buildInfo.description)</td><td>$($buildInfo.releaseDate)</td><td>$($buildInfo.date)</td><td><span class='version-badge win11-25h2'>Windows 11 25H2</span></td></tr>"
                        }
                    } else {
                        # Fallback voor hoofdbuild als er geen specifieke builds zijn
                        $kbEntries += "<tr class='build-row'><td>$baseBuild.xxxx</td><td>$($kbInfo.kb)</td><td>$($kbInfo.title)</td><td>$($kbInfo.releaseDate)</td><td>$($kbInfo.date)</td><td><span class='version-badge win11-25h2'>Windows 11 25H2</span></td></tr>"
                    }
                }
            }
            
            # Verwerk Windows 11 24H2 mappings
            if ($KBMappingForHTML.Data.mappings.windows11_24h2) {
                foreach ($baseBuild in ($KBMappingForHTML.Data.mappings.windows11_24h2.PSObject.Properties.Name | Sort-Object -Descending)) {
                    $kbInfo = $KBMappingForHTML.Data.mappings.windows11_24h2.$baseBuild
                    
                    # Specifieke builds (indien beschikbaar)
                    if ($kbInfo.builds) {
                        foreach ($specificBuild in ($kbInfo.builds.PSObject.Properties.Name | Sort-Object -Descending)) {
                            $buildInfo = $kbInfo.builds.$specificBuild
                            $kbEntries += "<tr class='build-row'><td>$specificBuild</td><td>$($buildInfo.kb)</td><td>$($buildInfo.description)</td><td>$($buildInfo.releaseDate)</td><td>$($buildInfo.date)</td><td><span class='version-badge win11-24h2'>Windows 11 24H2</span></td></tr>"
                        }
                    } else {
                        # Fallback voor hoofdbuild als er geen specifieke builds zijn
                        $kbEntries += "<tr class='build-row'><td>$baseBuild.xxxx</td><td>$($kbInfo.kb)</td><td>$($kbInfo.title)</td><td>$($kbInfo.releaseDate)</td><td>$($kbInfo.date)</td><td><span class='version-badge win11-24h2'>Windows 11 24H2</span></td></tr>"
                    }
                }
            }
            
            # Verwerk Windows 11 22H2/23H2 mappings
            if ($KBMappingForHTML.Data.mappings.windows11_22h2) {
                foreach ($baseBuild in ($KBMappingForHTML.Data.mappings.windows11_22h2.PSObject.Properties.Name | Sort-Object -Descending)) {
                    $kbInfo = $KBMappingForHTML.Data.mappings.windows11_22h2.$baseBuild
                    
                    # Specifieke builds (indien beschikbaar)
                    if ($kbInfo.builds) {
                        foreach ($specificBuild in ($kbInfo.builds.PSObject.Properties.Name | Sort-Object -Descending)) {
                            $buildInfo = $kbInfo.builds.$specificBuild
                            $kbEntries += "<tr class='build-row'><td>$specificBuild</td><td>$($buildInfo.kb)</td><td>$($buildInfo.description)</td><td>$($buildInfo.releaseDate)</td><td>$($buildInfo.date)</td><td><span class='version-badge win11-22h2'>Windows 11 $($kbInfo.version)</span></td></tr>"
                        }
                    } else {
                        # Fallback voor hoofdbuild als er geen specifieke builds zijn
                        $kbEntries += "<tr class='build-row'><td>$baseBuild.xxxx</td><td>$($kbInfo.kb)</td><td>$($kbInfo.title)</td><td>$($kbInfo.releaseDate)</td><td>$($kbInfo.date)</td><td><span class='version-badge win11-22h2'>Windows 11 $($kbInfo.version)</span></td></tr>"
                    }
                }
            }
            
            # Verwerk Windows 10 mappings
            if ($KBMappingForHTML.Data.mappings.windows10) {
                foreach ($baseBuild in ($KBMappingForHTML.Data.mappings.windows10.PSObject.Properties.Name | Sort-Object -Descending)) {
                    $kbInfo = $KBMappingForHTML.Data.mappings.windows10.$baseBuild
                    
                    # Specifieke builds (indien beschikbaar)
                    if ($kbInfo.builds) {
                        foreach ($specificBuild in ($kbInfo.builds.PSObject.Properties.Name | Sort-Object -Descending)) {
                            $buildInfo = $kbInfo.builds.$specificBuild
                            $kbEntries += "<tr class='build-row'><td>$specificBuild</td><td>$($buildInfo.kb)</td><td>$($buildInfo.description)</td><td>$($buildInfo.releaseDate)</td><td>$($buildInfo.date)</td><td><span class='version-badge win10'>Windows 10 $($kbInfo.version)</span></td></tr>"
                        }
                    } else {
                        # Fallback voor hoofdbuild als er geen specifieke builds zijn
                        $kbEntries += "<tr class='build-row'><td>$baseBuild.xxxx</td><td>$($kbInfo.kb)</td><td>$($kbInfo.title)</td><td>$($kbInfo.releaseDate)</td><td>$($kbInfo.date)</td><td><span class='version-badge win10'>Windows 10 $($kbInfo.version)</span></td></tr>"
                    }
                }
            }
            
            # Verwerk Historical mappings
            if ($KBMappingForHTML.Data.mappings.historical) {
                foreach ($year in ($KBMappingForHTML.Data.mappings.historical.PSObject.Properties.Name | Sort-Object -Descending)) {
                    foreach ($build in ($KBMappingForHTML.Data.mappings.historical.$year.PSObject.Properties.Name | Sort-Object -Descending)) {
                        $kbInfo = $KBMappingForHTML.Data.mappings.historical.$year.$build
                        $kbEntries += "<tr class='build-row'><td>$build</td><td>$($kbInfo.kb)</td><td>$($kbInfo.title)</td><td>$($kbInfo.date)</td><td>$year</td><td><span class='version-badge historical'>Historical</span></td></tr>"
                    }
                }
            }
            
            "$exportButtonsHtml
            <table id='kbMappingTable' class='display' style='width:100%'>
                <thead>
                    <tr>
                        <th>Build Number</th>
                        <th>KB Number</th>
                        <th>Update Title/Description</th>
                        <th>Release Date</th>
                        <th>Period</th>
                        <th>OS Version</th>
                    </tr>
                </thead>
                <tbody>
                    $($kbEntries -join "`n")
                </tbody>
            </table>"
        } else {
            "<p style='color:red; font-style:italic;'>KB Mapping database kon niet worden geladen. Controleer de configuratie en internetverbinding.</p>"
        })
    </div>
    
    <!-- Office Versions Tab -->
    <div id="OfficeVersions" class="tabcontent" style="display:none">
        <h2>Office Version Mapping Overview</h2>
        <div class="kb-status-box">
            <strong>Database Status:</strong> $(if ($OfficeMappingForHTML.Success) { "<span style='color:green;'>✓ Beschikbaar</span>" } else { "<span style='color:red;'>✗ Niet beschikbaar</span>" })<br>
            <strong>Bron Methode:</strong> $($OfficeMappingForHTML.Method)<br>
            <strong>Laatste Update:</strong> $($OfficeMappingForHTML.LastUpdated)<br>
            <strong>Versie:</strong> $($OfficeMappingForHTML.Version)
            $(if (-not $OfficeMappingForHTML.Success) { "<br><strong>Fout:</strong> <span style='color:red;'>$($OfficeMappingForHTML.Error)</span>" })
        </div>
        
        $(if ($OfficeMappingForHTML.Success -and $OfficeMappingForHTML.Data) {
            $officeEntries = @()
            
            # Export buttons HTML
            $officeExportButtonsHtml = "<div class='export-buttons' style='margin-bottom: 10px;'>
                <button onclick=`"exportTableToCSV('officeVersionTable', 'Office_Versions-full.csv', false)`">Exporteren volledige tabel</button>
                <button onclick=`"exportTableToCSV('officeVersionTable', 'Office_Versions-filtered.csv', true)`">Exporteren gefilterde rijen</button>
            </div>"
            
            # Verwerk Current Channel
            if ($OfficeMappingForHTML.Data.update_channels.'Current Channel') {
                $channelData = $OfficeMappingForHTML.Data.update_channels.'Current Channel'
                $officeEntries += "<tr class='build-row'><td>Current Channel</td><td>$($channelData.current_version)</td><td>$($channelData.current_build)</td><td>$($channelData.release_date)</td><td>$($channelData.update_frequency)</td><td><span class='version-badge' style='background:#28a745;color:white;'>Latest Features</span></td></tr>"
            }
            
            # Verwerk Monthly Enterprise Channel
            if ($OfficeMappingForHTML.Data.update_channels.'Monthly Enterprise Channel' -and $OfficeMappingForHTML.Data.update_channels.'Monthly Enterprise Channel'.versions) {
                foreach ($version in $OfficeMappingForHTML.Data.update_channels.'Monthly Enterprise Channel'.versions) {
                    $eolStatus = ""
                    if ($version.end_of_support) {
                        $eolDate = [DateTime]::Parse($version.end_of_support)
                        $daysUntilEOL = ($eolDate - (Get-Date)).Days
                        if ($daysUntilEOL -lt 0) {
                            $eolStatus = "<span class='version-badge' style='background:#dc3545;color:white;'>EOL $($version.end_of_support)</span>"
                        } elseif ($daysUntilEOL -lt 30) {
                            $eolStatus = "<span class='version-badge' style='background:#ffc107;color:#212529;'>EOL Soon ($daysUntilEOL days)</span>"
                        } else {
                            $eolStatus = "<span class='version-badge' style='background:#17a2b8;color:white;'>Supported</span>"
                        }
                    }
                    $noteText = if ($version.note) { " <em>($($version.note))</em>" } else { "" }
                    $officeEntries += "<tr class='build-row'><td>Monthly Enterprise$noteText</td><td>$($version.version)</td><td>$($version.build)</td><td>$($version.release_date)</td><td>Maandelijks</td><td>$eolStatus</td></tr>"
                }
            }
            
            # Verwerk Semi-Annual Enterprise Channel (Preview)
            if ($OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel (Preview)') {
                $channelData = $OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel (Preview)'
                $eolDate = [DateTime]::Parse($channelData.end_of_support)
                $daysUntilEOL = ($eolDate - (Get-Date)).Days
                if ($daysUntilEOL -lt 0) {
                    $eolStatus = "<span class='version-badge' style='background:#dc3545;color:white;'>EOL $($channelData.end_of_support)</span>"
                } elseif ($daysUntilEOL -lt 30) {
                    $eolStatus = "<span class='version-badge' style='background:#ffc107;color:#212529;'>EOL Soon ($daysUntilEOL days)</span>"
                } else {
                    $eolStatus = "<span class='version-badge' style='background:#6f42c1;color:white;'>Preview</span>"
                }
                $officeEntries += "<tr class='build-row'><td>Semi-Annual (Preview)</td><td>$($channelData.current_version)</td><td>$($channelData.current_build)</td><td>$($channelData.release_date)</td><td>2x per jaar</td><td>$eolStatus</td></tr>"
            }
            
            # Verwerk Semi-Annual Enterprise Channel
            if ($OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel' -and $OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel'.versions) {
                foreach ($version in $OfficeMappingForHTML.Data.update_channels.'Semi-Annual Enterprise Channel'.versions) {
                    $eolDate = [DateTime]::Parse($version.end_of_support)
                    $daysUntilEOL = ($eolDate - (Get-Date)).Days
                    if ($daysUntilEOL -lt 0) {
                        $eolStatus = "<span class='version-badge' style='background:#dc3545;color:white;'>EOL $($version.end_of_support)</span>"
                    } elseif ($daysUntilEOL -lt 30) {
                        $eolStatus = "<span class='version-badge' style='background:#ffc107;color:#212529;'>EOL Soon ($daysUntilEOL days)</span>"
                    } else {
                        $eolStatus = "<span class='version-badge' style='background:#28a745;color:white;'>Stable/LTS</span>"
                    }
                    $officeEntries += "<tr class='build-row'><td>Semi-Annual Enterprise</td><td>$($version.version)</td><td>$($version.build)</td><td>$($version.release_date)</td><td>2x per jaar</td><td>$eolStatus</td></tr>"
                }
            }
            
            "$officeExportButtonsHtml
            <table id='officeVersionTable' class='display' style='width:100%'>
                <thead>
                    <tr>
                        <th>Update Channel</th>
                        <th>Version</th>
                        <th>Build Number</th>
                        <th>Release Date</th>
                        <th>Update Frequency</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    $($officeEntries -join "`n")
                </tbody>
            </table>"
        } else {
            "<p style='color:red; font-style:italic;'>Office Version Mapping kon niet worden geladen. Controleer de configuratie.</p>"
        })
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

    // Functie om grafiek te resetten naar alle klanten
    function showAllCustomersChart() {
        chart.data.labels = [$ChartLabelsString];
        chart.data.datasets = [$ChartDatasets];
        chart.options.plugins.title.text = 'Alle klanten';
        chart.update();
    }

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

    // Functie voor KB Mapping Database tab
    function showKBMapping() {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        document.getElementById("KBMapping").style.display = "block";
        document.getElementsByClassName("tablinks")[2].className += " active";
        
        // Initialiseer DataTable voor KB Mapping (zonder groep headers)
        if (typeof initializeDataTable === 'function') {
            if (`$.fn.DataTable.isDataTable('#kbMappingTable')) {
                `$('#kbMappingTable').DataTable().destroy();
            }
            `$('#kbMappingTable').DataTable({
                "responsive": true,
                "pageLength": 25,
                "order": [[ 5, "desc" ], [ 0, "desc" ]], // Sorteer op Windows versie, dan build
                "language": {
                    "search": "Zoeken:",
                    "lengthMenu": "Toon _MENU_ regels",
                    "info": "Toon _START_ tot _END_ van _TOTAL_ regels",
                    "infoEmpty": "Geen gegevens beschikbaar",
                    "infoFiltered": "(gefilterd uit _MAX_ totaal regels)",
                    "paginate": {
                        "first": "Eerste",
                        "last": "Laatste",
                        "next": "Volgende", 
                        "previous": "Vorige"
                    },
                    "emptyTable": "Geen gegevens beschikbaar in de tabel"
                }
            });
        }
        
        // Reset grafiek naar alle klanten bij wisselen naar KB Mapping
        showAllCustomersChart();
    }
    
    // Functie voor Office Versions tab
    function showOfficeVersions() {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }
        document.getElementById("OfficeVersions").style.display = "block";
        document.getElementsByClassName("tablinks")[3].className += " active";
        
        // Initialiseer DataTable voor Office Versions
        if (typeof initializeDataTable === 'function') {
            if (`$.fn.DataTable.isDataTable('#officeVersionTable')) {
                `$('#officeVersionTable').DataTable().destroy();
            }
            `$('#officeVersionTable').DataTable({
                "responsive": true,
                "pageLength": 25,
                "order": [[ 2, "desc" ]], // Sorteer op build number (meest recent eerst)
                "language": {
                    "search": "Zoeken:",
                    "lengthMenu": "Toon _MENU_ regels",
                    "info": "Toon _START_ tot _END_ van _TOTAL_ regels",
                    "infoEmpty": "Geen gegevens beschikbaar",
                    "infoFiltered": "(gefilterd uit _MAX_ totaal regels)",
                    "paginate": {
                        "first": "Eerste",
                        "last": "Laatste",
                        "next": "Volgende", 
                        "previous": "Vorige"
                    },
                    "emptyTable": "Geen gegevens beschikbaar in de tabel"
                }
            });
        }
        
        // Reset grafiek naar alle klanten bij wisselen naar Office Versions
        showAllCustomersChart();
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
        showAllCustomersChart();
    }

    // Initialiseer de pagina bij het laden
    window.onload = function() {
        showAllCustomers();
    }
    
    // Direct uitvoeren na script load als backup
    document.addEventListener('DOMContentLoaded', function() {
        // Wacht even totdat alles geladen is
        setTimeout(function() {
            if (typeof showAllCustomers === 'function') {
                showAllCustomers();
            }
        }, 100);
    });
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
        Start-Process $HtmlPath
        Write-Host "Rapport geopend in standaard webbrowser." -ForegroundColor Green
    } catch {
        Write-Warning "Kon rapport niet automatisch openen: $_"
    }
} else {
    Write-Host "Automatisch openen van rapport is uitgeschakeld in configuratie." -ForegroundColor Yellow
    Write-Host "U kunt het rapport handmatig openen via: $HtmlPath" -ForegroundColor Cyan
}

Write-Host "`nScript voltooid! Alle rapporten zijn gegenereerd en beschikbaar in de exports directory." -ForegroundColor Green

# === BACKUP SECTIE ===
Write-Host "`n=== BACKUP PROCES ===" -ForegroundColor Magenta

if ($config.backup.enableExportBackup -eq $true -or $config.backup.enableArchiveBackup -eq $true -or $config.backup.enableConfigBackup -eq $true) {
    # Maak backup directories aan
    $BackupBaseDir = ".\backup"
    $BackupExportDir = "$BackupBaseDir\export_backup"
    $BackupArchiveDir = "$BackupBaseDir\archive_backup"
    $BackupConfigDir = "$BackupBaseDir\config_backup"
    
    @($BackupBaseDir, $BackupExportDir, $BackupArchiveDir, $BackupConfigDir) | ForEach-Object {
        if (-not (Test-Path $_)) { New-Item -Path $_ -ItemType Directory -Force | Out-Null }
    }

    # Backup van exports
    if ($config.backup.enableExportBackup -eq $true -and (Test-Path ".\$($config.exportDirectory)")) {
        $ExportZipPath = "$BackupExportDir\export-$(Get-Date -Format 'yyyyMMdd').zip"
        Compress-Archive -Path ".\$($config.exportDirectory)\*" -DestinationPath $ExportZipPath -Force
        
        # Behoud alleen de laatste X backups
        $zips = Get-ChildItem -Path $BackupExportDir -Filter "export-*.zip" | Sort-Object LastWriteTime -Descending
        if ($zips.Count -gt $config.backup.exportBackupRetention) {
            $zipsToRemove = $zips | Select-Object -Skip $config.backup.exportBackupRetention
            foreach ($z in $zipsToRemove) {
                Write-Host "Verwijder oude export backup: $($z.FullName)" -ForegroundColor Yellow
                Remove-Item $z.FullName -Force
            }
        }
        Write-Host "Backup gemaakt: $ExportZipPath" -ForegroundColor Green
    }

    # Backup van archive
    if ($config.backup.enableArchiveBackup -eq $true -and (Test-Path ".\archive")) {
        $ArchiveZipPath = "$BackupArchiveDir\archive-$(Get-Date -Format 'yyyyMMdd').zip"
        Compress-Archive -Path ".\archive\*" -DestinationPath $ArchiveZipPath -Force
        
        # Behoud alleen de laatste X backups
        $zips = Get-ChildItem -Path $BackupArchiveDir -Filter "archive-*.zip" | Sort-Object LastWriteTime -Descending
        if ($zips.Count -gt $config.backup.archiveBackupRetention) {
            $zipsToRemove = $zips | Select-Object -Skip $config.backup.archiveBackupRetention
            foreach ($z in $zipsToRemove) {
                Write-Host "Verwijder oude archive backup: $($z.FullName)" -ForegroundColor Yellow
                Remove-Item $z.FullName -Force
            }
        }
        Write-Host "Backup gemaakt: $ArchiveZipPath" -ForegroundColor Green
    }

    # Backup van configuratie bestanden
    if ($config.backup.enableConfigBackup -eq $true) {
        $ConfigFiles = @("config.json", "credentials.json", "_config.json", "_credentials.json", "kb-mapping.json")
        $ConfigFilesToBackup = $ConfigFiles | Where-Object { Test-Path $_ }
        
        if ($ConfigFilesToBackup.Count -gt 0) {
            $ConfigZipPath = "$BackupConfigDir\configcreds-$(Get-Date -Format 'yyyyMMdd').zip"
            Compress-Archive -Path $ConfigFilesToBackup -DestinationPath $ConfigZipPath -Force
            
            # Behoud alleen de laatste X backups
            $zips = Get-ChildItem -Path $BackupConfigDir -Filter "configcreds-*.zip" | Sort-Object LastWriteTime -Descending
            if ($zips.Count -gt $config.backup.configBackupRetention) {
                $zipsToRemove = $zips | Select-Object -Skip $config.backup.configBackupRetention
                foreach ($z in $zipsToRemove) {
                    Write-Host "Verwijder oude config backup: $($z.FullName)" -ForegroundColor Yellow
                    Remove-Item $z.FullName -Force
                }
            }
            Write-Host "Backup gemaakt: $ConfigZipPath" -ForegroundColor Green
        }
    }
}

Write-Host "Backups voltooid en opgeslagen in de backup directory." -ForegroundColor Cyan
    
