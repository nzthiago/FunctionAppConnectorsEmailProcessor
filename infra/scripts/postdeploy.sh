#!/bin/bash
set -e

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
WHITE='\033[1;37m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Creating AI Gateway trigger config...${NC}"

# Get outputs from azd
outputs=$(azd env get-values --output json)

if command -v jq &> /dev/null; then
    subscriptionId=$(echo "$outputs" | jq -r '.AZURE_SUBSCRIPTION_ID')
    resourceGroupName=$(echo "$outputs" | jq -r '.resourceGroupName')
    aiGatewayName=$(echo "$outputs" | jq -r '.aiGatewayName')
    aiGatewayConnectionName=$(echo "$outputs" | jq -r '.aiGatewayConnectionName')
    functionAppName=$(echo "$outputs" | jq -r '.functionAppName')
    functionAppDefaultHostname=$(echo "$outputs" | jq -r '.functionAppDefaultHostname')
    office365FunctionName=$(echo "$outputs" | jq -r '.office365FunctionName')
else
    echo -e "${RED}Error: jq is required for this script. Please install jq.${NC}"
    exit 1
fi

# Fetch the connector extension system key
echo -e "${CYAN}Fetching connector extension key for ${functionAppName}...${NC}"
connectorExtensionKey=$(az functionapp keys list -g "${resourceGroupName}" -n "${functionAppName}" --query "systemKeys.connector_extension" -o tsv)

triggerName="${aiGatewayConnectionName}-trigger"
callbackUrl="https://${functionAppName}.azurewebsites.net/runtime/webhooks/connector?functionName=${office365FunctionName}&code=${connectorExtensionKey}"

apiUrl="https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/aigateways/${aiGatewayName}/triggerconfigs/${triggerName}?api-version=2026-03-01-preview"

body=$(cat <<EOF
{
  "properties": {
    "description": "Office 365 Outlook trigger config",
    "connectionDetails": {
      "connectorName": "office365",
      "connectionName": "${aiGatewayConnectionName}"
    },
    "operationName": "OnNewEmailV3",
    "parameters": [
      {
        "name": "folderPath",
        "value": "Inbox"
      }
    ],
    "notificationDetails": {
      "callbackUrl": "${callbackUrl}"
    }
  }
}
EOF
)

echo -e "${CYAN}  API URL: ${apiUrl}${NC}"
echo -e "${CYAN}  Callback URL: ${callbackUrl}${NC}"

az rest --method PUT --url "${apiUrl}" --body "${body}"

echo -e "${GREEN}✅ AI Gateway trigger config created successfully!${NC}"

echo ""
echo -e "${YELLOW}╔══════════════════════════════════════════════════════════════════════╗${NC}"
echo -e "${YELLOW}║  ⚠️  IMPORTANT: Authorize the Office 365 Connection                 ║${NC}"
echo -e "${YELLOW}╠══════════════════════════════════════════════════════════════════════╣${NC}"
echo -e "${YELLOW}║                                                                      ║${NC}"
echo -e "${YELLOW}║  Before testing, you must authorize the Office 365 connector:        ║${NC}"
echo -e "${YELLOW}║                                                                      ║${NC}"
echo -e "${YELLOW}║  1. Open the Azure Portal: https://portal.azure.com                  ║${NC}"
echo -e "${YELLOW}║  2. Navigate to Resource Group: ${resourceGroupName}${NC}"
echo -e "${YELLOW}║  3. Open the AI Gateway resource: ${aiGatewayName}${NC}"
echo -e "${YELLOW}║  4. Go to Connections → select the Office 365 connection             ║${NC}"
echo -e "${YELLOW}║  5. Click 'Authorize' and sign in with your Office 365 account       ║${NC}"
echo -e "${YELLOW}║                                                                      ║${NC}"
echo -e "${YELLOW}║  The trigger will NOT fire until the connection is authorized.        ║${NC}"
echo -e "${YELLOW}╚══════════════════════════════════════════════════════════════════════╝${NC}"
echo ""
