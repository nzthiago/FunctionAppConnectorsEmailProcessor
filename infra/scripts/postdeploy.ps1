Write-Host "Creating AI Gateway trigger config..." -ForegroundColor Yellow

# Get outputs from azd
$outputs = azd env get-values --output json | ConvertFrom-Json

$subscriptionId = (az account show --query id -o tsv)
$resourceGroupName = $outputs.resourceGroupName
$aiGatewayName = $outputs.aiGatewayName
$aiGatewayConnectionName = $outputs.aiGatewayConnectionName
$functionAppName = $outputs.functionAppName
$functionAppDefaultHostname = $outputs.functionAppDefaultHostname
$office365FunctionName = $outputs.office365FunctionName

# Fetch the function key
Write-Host "Fetching function key for $office365FunctionName..." -ForegroundColor Cyan
$functionKey = (az functionapp function keys list --function-name $office365FunctionName --name $functionAppName --resource-group $resourceGroupName --query "default" -o tsv)

$triggerName = "$aiGatewayConnectionName-trigger"
$callbackUrl = "https://$functionAppDefaultHostname/api/$office365FunctionName?code=$functionKey"

$apiUrl = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Web/aigateways/$aiGatewayName/triggerconfigs/${triggerName}?api-version=2026-03-01-preview"

$body = @{
  properties = @{
    description = "Office 365 Outlook trigger config"
    connectionDetails = @{
      connectorName = "office365"
      connectionName = $aiGatewayConnectionName
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
} | ConvertTo-Json -Depth 5 -Compress

Write-Host "  API URL: $apiUrl" -ForegroundColor Cyan
Write-Host "  Callback URL: $callbackUrl" -ForegroundColor Cyan

az rest --method PUT --url $apiUrl --body $body

if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to create AI Gateway trigger config." -ForegroundColor Red
    exit 1
}

Write-Host "✅ AI Gateway trigger config created successfully!" -ForegroundColor Green

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║  ⚠️  IMPORTANT: Authorize the Office 365 Connection                 ║" -ForegroundColor Yellow
Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  Before testing, you must authorize the Office 365 connector:        ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  1. Open the Azure Portal: https://portal.azure.com                  ║" -ForegroundColor Yellow
Write-Host "║  2. Navigate to Resource Group: $resourceGroupName" -ForegroundColor Yellow
Write-Host "║  3. Open the AI Gateway resource: $aiGatewayName" -ForegroundColor Yellow
Write-Host "║  4. Go to Connections -> select the Office 365 connection             ║" -ForegroundColor Yellow
Write-Host "║  5. Click 'Authorize' and sign in with your Office 365 account       ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  The trigger will NOT fire until the connection is authorized.        ║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""
