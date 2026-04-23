# Azure Functions Connectors Sample

This sample demonstrates how to use **Azure Functions** with **Connector Gateway connectors** to react to events from external services — in this case, receiving notifications when new emails arrive in an Office 365 Outlook inbox.

## Architecture

- **Azure Functions (Flex Consumption)** — A .NET 10 isolated worker function app that receives HTTP callbacks when new emails arrive.
- **Connector Gateway** — Manages the Office 365 Outlook connector and trigger configuration, routing events to the function app.
- **Office 365 Outlook Connector** — Monitors an inbox for new emails and triggers a notification to the function.

## Prerequisites

- [Azure Developer CLI (azd)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli)
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [jq](https://jqlang.github.io/jq/) (required by the post-deploy script on Linux/macOS)
- An Azure subscription
- An Office 365 account (for the email connector)

## Getting Started

### 1. Clone Required Repositories

This project references two companion libraries via local project references (NuGet packages are not yet available). Clone all three repositories into the **same parent directory**:

```bash
git clone <url>/FunctionAppConnectorsEmailProcessor
git clone <url>/azure-functions-connector-extension
git clone <url>/azure-logicapps-connector-sdk
```

Your folder structure should look like:

```
connectors/
├── FunctionAppConnectorsEmailProcessor/
├── azure-functions-connector-extension/
└── azure-logicapps-connector-sdk/
```

### 2. Deploy to Azure

```bash
cd FunctionAppConnectorsEmailProcessor
azd up
```

This provisions all infrastructure (Function App, Connector Gateway, Storage, Application Insights) and deploys the function code. After deployment, a post-deploy script automatically creates the Connector Gateway trigger configuration.

### 3. Authorize the Office 365 Connection

> **⚠️ Important:** After deployment, you **must** authorize the Office 365 connector connection before the trigger will work. The connection is created in a disabled state and requires your OAuth consent.

1. Open the [Azure Portal](https://portal.azure.com).
2. Navigate to the **Resource Group** created by the deployment.
3. Open the **Connector Gateway** resource.
4. Go to **Connections** and select the Office 365 connection.
5. Click **Authorize** and sign in with the Office 365 account whose inbox you want to monitor.
6. Confirm the consent prompt to grant the connector access to your mailbox.

Until this step is completed, the trigger will not fire and the function will not receive email notifications.

### 4. Test the Solution

Once the connection is authorized, the trigger monitors the configured inbox folder (`Inbox` by default). It takes a few minutes for the connection and trigger to start. Send an email to the authorized account and the `OnNewImportantEmailReceived` function will be invoked via an HTTP POST callback.

You can also manually test the function endpoint using the [test.http](test.http) file (update the URL and function key to match your deployment).

## Project Structure

| Path | Description |
|---|---|
| `function-app/` | Azure Functions application (.NET 10, isolated worker) |
| `function-app/ProcessEmail.cs` | Function triggered when a new email arrives |
| `function-app/Program.cs` | Host builder configuration |
| `infra/main.bicep` | Main Bicep template for all Azure resources |
| `infra/connectorGateway.bicep` | Connector Gateway and connection resource definitions |
| `infra/scripts/postdeploy.sh` | Post-deploy script (Linux/macOS) — creates the trigger config |
| `infra/scripts/postdeploy.ps1` | Post-deploy script (Windows) — creates the trigger config |
| `azure.yaml` | Azure Developer CLI project configuration |
| `test.http` | Sample HTTP request for manual testing |

## Resources

- [Azure Functions documentation](https://learn.microsoft.com/azure/azure-functions/)
- [Azure Functions Flex Consumption plan](https://learn.microsoft.com/azure/azure-functions/flex-consumption-plan)
- [Azure Developer CLI (azd)](https://learn.microsoft.com/azure/developer/azure-developer-cli/)
