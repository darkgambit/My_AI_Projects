param(
    [Parameter(Mandatory=$true)]
    [string]$WorkflowFile
)

# Ensure env vars are set
if (-not $env:N8N_API_URL) {
    $env:N8N_API_URL = "https://n8n-d7yx.srv1842740.hstgr.cloud"
}

if (-not $env:N8N_API_KEY) {
    Write-Error "N8N_API_KEY is not set!"
    exit 1
}

if (-not (Test-Path $WorkflowFile)) {
    Write-Error "Workflow file not found: $WorkflowFile"
    exit 1
}

# Use curl.exe for more reliable HTTP handling
$headers = "X-N8N-API-KEY: $env:N8N_API_KEY"
$contentType = "Content-Type: application/json"
$uri = "$env:N8N_API_URL/api/v1/workflows"

Write-Host "Importing: $WorkflowFile"
Write-Host "To: $uri"

& curl.exe -X POST `
    -H $headers `
    -H $contentType `
    -d "@$WorkflowFile" `
    $uri

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Done!"
} else {
    Write-Host "❌ Import failed with exit code: $LASTEXITCODE"
}
