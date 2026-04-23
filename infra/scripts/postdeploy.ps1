Write-Host "Post-deployment configuration..." -ForegroundColor Yellow

# Get outputs from azd
$outputs = azd env get-values --output json | ConvertFrom-Json

$subscriptionId = (az account show --query id -o tsv)
$resourceGroupName = $outputs.resourceGroupName
$connectorGatewayName = $outputs.connectorGatewayName
$connectorGatewayConnectionName = $outputs.connectorGatewayConnectionName
$functionAppName = $outputs.functionAppName
$functionAppDefaultHostname = $outputs.functionAppDefaultHostname
$office365FunctionName = $outputs.office365FunctionName

# --- Create Connector Gateway trigger config ---
Write-Host "Creating Connector Gateway trigger config..." -ForegroundColor Yellow

# Fetch the connector extension system key
Write-Host "Fetching connector extension key for $functionAppName..." -ForegroundColor Cyan
$connectorExtensionKey = (az functionapp keys list -g $resourceGroupName -n $functionAppName --query "systemKeys.connector_extension" -o tsv)

$triggerName = "$connectorGatewayConnectionName-trigger"
$callbackUrl = "https://$functionAppName.azurewebsites.net/runtime/webhooks/connector?functionName=$office365FunctionName&code=$connectorExtensionKey"

$apiUrl = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Web/connectorGateways/$connectorGatewayName/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

$body = @{
  properties = @{
    description = "Office 365 Outlook trigger config"
    connectionDetails = @{
      connectorName = "office365"
      connectionName = $connectorGatewayConnectionName
    }
    operationName = "OnNewEmailV3"
    parameters = @(
      @{ name = "folderPath"; value = "Inbox" }
      @{ name = "fetchOnlyWithAttachment"; value = "false" }
      @{ name = "includeAttachments"; value = "false" }
    )
    notificationDetails = @{
      callbackUrl = $callbackUrl
    }
  }
} | ConvertTo-Json -Depth 5

$bodyFile = [System.IO.Path]::GetTempFileName()
$body | Out-File -FilePath $bodyFile -Encoding utf8

Write-Host "  API URL: $apiUrl" -ForegroundColor Cyan
Write-Host "  Callback URL: $callbackUrl" -ForegroundColor Cyan

az rest --method PUT --url $apiUrl --body "@$bodyFile" --headers "Content-Type=application/json"
Remove-Item $bodyFile -ErrorAction SilentlyContinue

if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to create Connector Gateway trigger config." -ForegroundColor Red
    exit 1
}

Write-Host "✅ Connector Gateway trigger config created successfully!" -ForegroundColor Green

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║  ⚠️  IMPORTANT: Authorize the Office 365 AND Teams Connections       ║" -ForegroundColor Yellow
Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  Before testing, you must authorize both connectors:                 ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  1. Open the Azure Portal: https://portal.azure.com                  ║" -ForegroundColor Yellow
Write-Host "║  2. Navigate to Resource Group: $resourceGroupName" -ForegroundColor Yellow
Write-Host "║  3. Open the Connector Gateway resource: $connectorGatewayName" -ForegroundColor Yellow
Write-Host "║  4. Go to Connections -> authorize the Office 365 connection          ║" -ForegroundColor Yellow
Write-Host "║  5. Go to Connections -> authorize the Teams connection               ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  The trigger will NOT fire until Office 365 connection is authorized. ║" -ForegroundColor Yellow
Write-Host "║  Teams notifications require the Teams connection to be authorized.   ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  After authorizing Teams, set TEAMS_CONNECTION_RUNTIME_URL,           ║" -ForegroundColor Yellow
Write-Host "║  TEAMS_TEAM_ID, and TEAMS_CHANNEL_ID in the Function App settings.   ║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""
