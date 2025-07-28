<#
.SYNOPSIS
Genereert een Windows Update rapportage voor meerdere tenants via Microsoft Graph.

.DESCRIPTION
Dit script haalt per tenant de ontbrekende Windows-updates op via de Microsoft Graph Threat Hunting API.
De resultaten worden geëxporteerd naar CSV-bestanden en een HTML-dashboard met filterbare tabellen en grafieken.

.BENODIGDHEDEN
- PowerShell 7+
- Microsoft Graph PowerShell SDK
- Een Azure AD App Registration per tenant met de juiste permissies

.GEBRUIK
1. Vul het credentials.json bestand met de juiste tenantgegevens.
2. Voer het script uit: .\get-windows-update-report.ps1
3. Bekijk de resultaten in de map 'exports'.

.AUTEUR
Maarten Schmeitz (info@scns.nl  | https://www.scns.nl)

.LASTEDIT
2025-07-28

.VERSIE
1.0.0
#>

#Import JSON
# Read the JSON file
$json = Get-Content -Path ".\credentials.json" -Raw

# Convert JSON to PowerShell object
$data = $json | ConvertFrom-Json

foreach ($cred in $data.LoginCredentials) {

    $ClientID = "$($cred.ClientID)"
    $Secret = "$($cred.Secret)"
    $TenantID = "$($cred.TenantID)"

    #Collect App Secret
    $Secret = ConvertTo-SecureString $Secret -AsPlainText -Force
    $ClientSecretCredential = New-Object System.Management.Automation.PSCredential -ArgumentList ($ClientID, $Secret)

    #Connect to Graph using Application Secret
    Connect-MgGraph -TenantId $TenantID -ClientSecretCredential $ClientSecretCredential

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
    $ExportPath = ".\exports\$DateStamp`_$($cred.customername)_Windows_Update_report_Overview.csv"
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
    $ExportPathUpdates = ".\exports\$DateStamp`_$($cred.customername)_Windows_Update_report_ByUpdate.csv"
    $MissingUpdateTable | Export-Csv -NoTypeInformation -Path $ExportPathUpdates

    #Disconnect from MG Graph
    Disconnect-MgGraph
}

# Verzamel alle Overview-bestanden
$OverviewFiles = Get-ChildItem -Path ".\exports" -Filter "*_Overview.csv" | Sort-Object Name

# Haal de totalen per dag per klant op
$CountsPerDayPerCustomer = @{}
$LatestDatePerCustomer = @{}
$LatestCsvPerCustomer = @{}
foreach ($file in $OverviewFiles) {
    $csv = Import-Csv $file.FullName
    $TotalCount = ($csv | Measure-Object -Property Count -Sum).Sum
    $parts = $file.Name -split "_"
    $Date = $parts[0]
    $Customer = $parts[1]
    if (-not $CountsPerDayPerCustomer.ContainsKey($Customer)) {
        $CountsPerDayPerCustomer[$Customer] = @()
    }
    $CountsPerDayPerCustomer[$Customer] += [PSCustomObject]@{
        Date = $Date
        TotalCount = $TotalCount
    }
    # Bepaal de laatste datum per klant
    if (-not $LatestDatePerCustomer.ContainsKey($Customer) -or ($Date -gt $LatestDatePerCustomer[$Customer])) {
        $LatestDatePerCustomer[$Customer] = $Date
        $LatestCsvPerCustomer[$Customer] = $csv
    }
}

# Genereer Chart.js datasets per klant
$ChartDatasets = ""
$ChartLabels = @()
foreach ($Customer in $CountsPerDayPerCustomer.Keys) {
    $Data = ($CountsPerDayPerCustomer[$Customer] | ForEach-Object { $_.TotalCount }) -join ","
    $Labels = ($CountsPerDayPerCustomer[$Customer] | ForEach-Object { "'$($_.Date)'" })
    if ($Labels.Count -gt $ChartLabels.Count) { $ChartLabels = $Labels }
    $Color = "rgb($(Get-Random -Minimum 0 -Maximum 255),$(Get-Random -Minimum 0 -Maximum 255),$(Get-Random -Minimum 0 -Maximum 255))"
    $ChartDatasets += @"
        {
            label: '$Customer',
            data: [$Data],
            borderColor: '$Color',
            backgroundColor: '$Color',
            fill: false,
            tension: 0.2
        },
"@
}
$ChartLabelsString = $ChartLabels -join ","

# Genereer tabbladen en tabellen voor alleen de laatste datum per klant
$CustomerTabs = ""
$CustomerTables = ""
foreach ($Customer in $LatestCsvPerCustomer.Keys) {
    $TableRows = ""
    foreach ($row in $LatestCsvPerCustomer[$Customer]) {
        $TableRows += "<tr><td>$($row.Device)</td><td>$($row.'Missing Updates')</td><td>$($row.Count)</td><td>$($row.LastSeen)</td><td>$($row.LoggedOnUsers)</td></tr>`n"
    }
    $CustomerTabs += "<button class='tablinks' onclick=""openCustomer(event, '$Customer')"">$Customer</button>"
    $CustomerTables += @"
    <div id="$Customer" class="tabcontent" style="display:none">
        <h2>Laatste overzicht voor $Customer ($($LatestDatePerCustomer[$Customer]))</h2>
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
    $("table.display").each(function(){
        var table = $(this).DataTable({
            "order": [[2, "desc"]],
            "language": {
                "url": "//cdn.datatables.net/plug-ins/1.13.6/i18n/nl-NL.json"
            }
        });
        // Kolomfilters toevoegen
        $(this).find('thead th').each(function (i) {
            var title = $(this).text();
            $(this).append('<br><input type="text" placeholder="Filter '+title+'" style="width:90%;font-size:12px;" />');
            $(this).find("input").on('keyup change', function () {
                if (table.column(i).search() !== this.value) {
                    table.column(i).search(this.value).draw();
                }
            });
        });
    });
    // Open eerste tab standaard
    $(".tablinks").first().click();
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
    </script>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        .container { max-width: 1200px; margin: auto; }
        canvas { background: #fff; }
        table.dataTable thead th { background: #eee; }
        .tab { overflow: hidden; border-bottom: 1px solid #ccc; }
        .tab button { background-color: #f1f1f1; float: left; border: none; outline: none; cursor: pointer; padding: 10px 20px; transition: 0.3s; }
        .tab button:hover { background-color: #ddd; }
        .tab button.active { background-color: #ccc; }
        .tabcontent { display: none; padding: 20px 0; }
    </style>
</head>
<body>
<div class="container">
    <h1>Windows Update Overview</h1>
    <p>Laatst uitgevoerd op: $LastRunDate</p>
    <h2>Totale Count per dag per klant</h2>
    <canvas id="countChart" height="100"></canvas>
    <div class="tab">
        $CustomerTabs
    </div>
    $CustomerTables
</div>
<script>
    // Chart.js
    const ctx = document.getElementById('countChart').getContext('2d');
    new Chart(ctx, {
        type: 'line',
        data: {
            labels: [$ChartLabelsString],
            datasets: [
                $ChartDatasets
            ]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: true } }
        }
    });

    // Tabs
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
    }
</script>
</body>
</html>
"@

$HtmlPath = ".\exports\Windows_Update_Overview.html"
Set-Content -Path $HtmlPath -Value $Html -Encoding UTF8