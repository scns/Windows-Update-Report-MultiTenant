# GitHub Labels Creation Script
# This script creates the labels defined in labels.yml manually

$labels = @(
    @{ name = "breaking-change"; color = "ee0701"; description = "A breaking change for existing users." },
    @{ name = "major-change"; color = "ee0701"; description = "A major change for all users." },
    @{ name = "bugfix"; color = "ee0701"; description = "Inconsistencies or issues which will cause a problem for users or implementors." },
    @{ name = "documentation"; color = "0052cc"; description = "Solely about the documentation of the project." },
    @{ name = "enhancement"; color = "1d76db"; description = "Enhancement of the code, not introducing new features." },
    @{ name = "refactor"; color = "1d76db"; description = "Improvement of existing code, not introducing new features." },
    @{ name = "performance"; color = "1d76db"; description = "Improvement of performance, not introducing new features." },
    @{ name = "new-feature"; color = "1d76db"; description = "New features for the project." },
    @{ name = "maintenance"; color = "fbca04"; description = "Generic maintenance tasks." },
    @{ name = "ci"; color = "fbca04"; description = "Work that improves the continue integration." },
    @{ name = "dependencies"; color = "0366d6"; description = "Work that touches dependencies." },
    @{ name = "security"; color = "ee0701"; description = "Marks a security issue that needs to be resolved asap." },
    @{ name = "sync"; color = "ededed"; description = "Sync operations." },
    @{ name = "minor"; color = "0075ca"; description = "Minor changes." },
    @{ name = "chore"; color = "fef2c0"; description = "Chore tasks." }
)

$owner = "scns"
$repo = "Windows-Update-Report-MultiTenant"

# You need a GitHub token with repo permissions
# Set it as an environment variable: $env:GITHUB_TOKEN = "your_token_here"

foreach ($label in $labels) {
    $body = @{
        name = $label.name
        color = $label.color
        description = $label.description
    } | ConvertTo-Json

    $headers = @{
        Authorization = "token $env:GITHUB_TOKEN"
        Accept = "application/vnd.github.v3+json"
    }

    try {
        $uri = "https://api.github.com/repos/$owner/$repo/labels"
        Write-Host "Creating label: $($label.name)" -ForegroundColor Green
        Invoke-RestMethod -Uri $uri -Method Post -Body $body -Headers $headers -ContentType "application/json"
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 422) {
            Write-Host "Label '$($label.name)' already exists, updating..." -ForegroundColor Yellow
            $updateUri = "https://api.github.com/repos/$owner/$repo/labels/$($label.name)"
            try {
                Invoke-RestMethod -Uri $updateUri -Method Patch -Body $body -Headers $headers -ContentType "application/json"
                Write-Host "Updated label: $($label.name)" -ForegroundColor Cyan
            }
            catch {
                Write-Host "Failed to update label: $($label.name) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Failed to create label: $($label.name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "Label creation/update completed!" -ForegroundColor Green
