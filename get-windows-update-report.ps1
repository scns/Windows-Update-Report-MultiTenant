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
| distinct DeviceName, RecommendedSecurityUpdate
| summarize MissingUpdates=make_set(RecommendedSecurityUpdate) by DeviceName
| extend Count = array_length(MissingUpdates)
| join kind=leftouter (
    DeviceInfo
    | summarize arg_max(Timestamp, *) by DeviceName
    | project DeviceName, LastSeen=Timestamp, LoggedOnUsers
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