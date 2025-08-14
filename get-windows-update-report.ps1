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

foreach ($cred in $data.LoginCredentials) {

    Write-Host "Verwerken van klant: $($cred.customername)" -ForegroundColor Cyan

    $ClientID = "$($cred.ClientID)"
    $Secret = "$($cred.Secret)"
    $TenantID = "$($cred.TenantID)"

    #Collect App Secret
    $Secret = ConvertTo-SecureString $Secret -AsPlainText -Force
    $ClientSecretCredential = New-Object System.Management.Automation.PSCredential -ArgumentList ($ClientID, $Secret)

    #Connect to Graph using Application Secret
    Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome | Out-Null

    #Create Query
    $Query = "DeviceTvmSoftwareVulnerabilities
    | distinct DeviceName, RecommendedSecurityUpdate, OSPlatform
    | where OSPlatform != 'Linux'
    | summarize MissingUpdates=make_set(RecommendedSecurityUpdate) by DeviceName
    | extend Count = array_length(MissingUpdates)
    | join kind=leftouter (
        DeviceInfo
        | summarize arg_max(Timestamp, *) by DeviceName
        | project DeviceName, LastSeen=Timestamp, LoggedOnUsers, OSPlatform
    ) on DeviceName
    "

    #Format Query as JSON
    $Body = @{
        Query = $Query
    } | ConvertTo-Json

    #Run the query
    $Result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/security/runHuntingQuery" -Body $body

    #Format the results into an array
    $ResultsTable = $Result.results | ForEach-Object {
        $MissingUpdates = $_.MissingUpdates -join ','
        [PSCustomObject]@{
            Device = $_.DeviceName
            "Missing Updates" = ($MissingUpdates)
            "Count" = ($_.Count - 1)
            "LastSeen" = $_.LastSeen
            "LoggedOnUsers" = if ($_.LoggedOnUsers -is [System.Array]) { $_.LoggedOnUsers -join ', ' } else { $_.LoggedOnUsers }
        }
    }

    #Export the results
    $DateStamp = Get-Date -Format "yyyyMMdd"
    $ExportPath = ".\$($config.exportDirectory)\$DateStamp`_$($cred.customername)_Windows_Update_report_Overview.csv"
    $ResultsTable | Export-Csv -NoTypeInformation -path $ExportPath

    # Create a table: Missing Update -> Devices
    $MissingUpdateTable = @()
    foreach ($row in $Result.results) {
        foreach ($update in $row.MissingUpdates) {
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
    
    Write-Host "Klant $($cred.customername) voltooid. CSV-bestanden gegenereerd." -ForegroundColor Green
}

# Voer archivering uit na alle exports
Move-OldExportsToArchive -ExportPath $ExportDir -ArchivePath $ArchiveDir -RetentionCount $config.exportRetentionCount

# Verzamel alle Overview-bestanden
$OverviewFiles = Get-ChildItem -Path ".\$($config.exportDirectory)" -Filter "*_Overview.csv" | Sort-Object Name

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
            $TableRows += "<tr><td>$($row.Device)</td><td>$($row.'Missing Updates')</td><td>$($row.Count)</td><td>$($row.LastSeen)</td><td>$($row.LoggedOnUsers)</td></tr>`n"
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
        <button onclick="exportTableToCSV('overviewTable_$Customer', '$Customer-full.csv', false)">Exporteren volledige tabel</button>
        <button onclick="exportTableToCSV('overviewTable_$Customer', '$Customer-filtered.csv', true)">Exporteren gefilterde rijen</button>
        <table id="overviewTable_$Customer" class="display" style="width:100%">
            <thead>
                <tr>
                    <th>Device</th>
                    <th>Missing Updates</th>
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
                $(this).append('<br><input type="text" placeholder="Filter '+title+'" style="width:90%;font-size:12px;" />');
                $(this).find("input").on('keyup change', function () {
                    if (table.column(i).search() !== this.value) {
                        table.column(i).search(this.value).draw();
                    }
                });
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
    <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script>
    $DataTablesScript
    // Dark mode toggle
    document.addEventListener('DOMContentLoaded', function() {
        var btn = document.getElementById('darkModeToggle');
        if (btn) {
            btn.addEventListener('click', function() {
                document.body.classList.toggle('darkmode');
                btn.textContent = document.body.classList.contains('darkmode') ? '☀️ Light mode' : '🌙 Dark mode';
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
    </style>
</head>
<body>
<div class="container">
    <h1>Windows Update Overview <button id="darkModeToggle" style="float:right;margin-left:20px;">🌙 Dark mode</button></h1>

    <p>Laatst uitgevoerd op: $LastRunDate</p>
    <h2>Totale Count per dag per klant</h2>
    <canvas id="countChart" height="100"></canvas>
    <div class="tab">
        <button class="tablinks" onclick="showAllCustomers()">Alle klanten</button>
        $CustomerTabs
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
        
        // Reset grafiek naar alle klanten
        chart.data.labels = [$ChartLabelsString];
        chart.data.datasets = [
            $ChartDatasets
        ];
        chart.options.plugins.title.text = 'Alle klanten';
        chart.update();
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