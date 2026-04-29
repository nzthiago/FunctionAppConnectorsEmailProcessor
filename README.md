# Azure Functions Connectors Sample

This sample demonstrates how to use **Azure Functions** with **Connector Gateway connectors** to react to events from external services. It listens for **high-importance** emails arriving in an Office 365 inbox, enriches the sender with directory data from **Microsoft Graph (Groups & Users)**, and posts a formatted notification to a **Microsoft Teams** channel.

## Architecture

![Architecture diagram](docs/architecture.png)

> Editable source: [docs/architecture.drawio](docs/architecture.drawio) (open with [draw.io](https://app.diagrams.net)).

- **Azure Functions (Flex Consumption)** — A .NET 10 isolated worker function app that receives HTTP callbacks from the Connector Gateway.
- **Connector Gateway** — Manages three connections (Office 365, Microsoft Graph, Teams) and the trigger configuration.
- **Office 365 Outlook Connector** — Monitors the inbox. **Filtering happens server-side** via the trigger config (`folderPath: Inbox`, `importance: High`), so the function is only invoked for high-importance emails. This avoids unnecessary function executions.
- **Microsoft Graph (Groups & Users) Connector** — For each triggered email, the function looks up the sender to enrich the Teams message with display name, job title, and department.
- **Teams Connector** — Posts the enriched notification card to a configured Teams channel.

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

### 3. Authorize the Connections

> **⚠️ Important:** After deployment, you **must** authorize all three connector connections in the Azure portal before the end-to-end flow will work. Each connection is created in a disabled state and requires OAuth consent.

1. Open the [Azure Portal](https://portal.azure.com).
2. Navigate to the **Resource Group** created by the deployment.
3. Open the **Connector Gateway** resource.
4. Go to **Connections** and authorize each of the three connections in turn:
   - **Office 365** — sign in with the account whose inbox you want to monitor (drives the trigger).
   - **Microsoft Graph (Groups & Users)** — sign in with an account that can read directory users (drives sender enrichment).
   - **Teams** — sign in with an account that can post to the target Teams channel.

Until all three connections are authorized, the trigger will not fire and/or notifications will fail.

### 4. Test the Solution

Once the connections are authorized, send a **high-importance** email to the authorized account. The Connector Gateway trigger filters by `importance: High` server-side, so only those emails invoke the function. The function will look up the sender via Microsoft Graph and post an enriched notification to the configured Teams channel.

You can also manually test the function endpoint using the [test.http](test.http) file (update the URL and function key to match your deployment).

## Project Structure

| Path | Description |
|---|---|
| `function-app/` | Azure Functions application (.NET 10, isolated worker) |
| `function-app/ProcessEmail.cs` | Function triggered for high-importance emails; enriches sender via Graph and posts to Teams |
| `function-app/Program.cs` | Host builder, registers Teams and Microsoft Graph connector clients |
| `infra/main.bicep` | Main Bicep template for all Azure resources |
| `infra/connectorGateway.bicep` | Connector Gateway plus Office 365, Microsoft Graph, and Teams connection resources |
| `infra/scripts/postdeploy.sh` | Post-deploy script (Linux/macOS) — creates the trigger config with `importance: High` filter |
| `infra/scripts/postdeploy.ps1` | Post-deploy script (Windows) — creates the trigger config with `importance: High` filter |
| `azure.yaml` | Azure Developer CLI project configuration |
| `test.http` | Sample HTTP request for manual testing |

## Resources

- [Azure Functions documentation](https://learn.microsoft.com/azure/azure-functions/)
- [Azure Functions Flex Consumption plan](https://learn.microsoft.com/azure/azure-functions/flex-consumption-plan)
- [Azure Developer CLI (azd)](https://learn.microsoft.com/azure/developer/azure-developer-cli/)
