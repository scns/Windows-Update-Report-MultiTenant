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
| extend Count = array_length(MissingUpdates)"

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
        Count = ($_.Count - 1)
    }
}

#Export the results
$DateStamp = Get-Date -Format "yyyyMMdd"
$ExportPath = ".\exports\$DateStamp`_$($cred.customername)_Windows_Update_report.csv"
$ResultsTable | Export-Csv -NoTypeInformation -path $ExportPath

#Disconnect from MG Graph
Disconnect-MgGraph
}