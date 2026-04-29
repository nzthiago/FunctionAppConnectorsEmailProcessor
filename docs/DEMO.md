# Demo guide — Important Email Triage with Connector Gateway

A two-part walkthrough:

1. **[Part 1 — Local debug](#part-1--local-debug-with-vs-code)** — show the code paths and decisions step-by-step in the VS Code debugger using `func start` + a synthetic payload from [test.http](../test.http).
2. **[Part 2 — Cloud walkthrough](#part-2--cloud-walkthrough)** — show the deployed resources in the Azure Portal, the Connector Gateway configuration, and trigger an end-to-end run from a real email.

> Architecture refresher: ![architecture](architecture.png)
> Source: [docs/architecture.drawio](architecture.drawio)

---

## Part 1 — Local debug with VS Code

### Prerequisites

- `azd up` already run and connections authorized (see [README](../README.md)).
- [Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite) running (the VS Code extension's "Start Azurite" command works great).
- `az login --tenant <your tenant>` (so `DefaultAzureCredential` can reach the Connector Gateway runtime URLs locally).
- A populated [function-app/local.settings.json](../function-app/local.settings.json) — values come from `azd env get-values` plus the `TEAMS_CONNECTION_RUNTIME_URL` / `GRAPH_CONNECTION_RUNTIME_URL` app settings on the deployed Function App.
- The C# Dev Kit + Azure Functions VS Code extensions.

### Launch the host under the debugger

1. Open the workspace in VS Code.
2. **Run and Debug** → "Attach to .NET Functions" (or just hit **F5** from the function-app folder; the Azure Functions extension scaffolds a launch config the first time).
3. Wait for the function listing to print:
   ```
   Functions:
     OnNewImportantEmailReceived: [POST] http://localhost:7071/api/OnNewImportantEmailReceived
   ```

### Set the test.http variables for local

At the top of [test.http](../test.http):

```
@functionAppUrl = http://localhost:7071
@httpPostCode =
```

### The breakpoint tour

Set these breakpoints in order, then click **Send Request** above the **first** POST in [test.http](../test.http) (the urgent one).

#### 🔴 BP 1 — Trigger entry: see the raw connector payload
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var emails = payload.Body?.Value ?? [];`

- Inspect `payload` in **Locals**. This is the exact JSON the Connector Gateway POSTs back, deserialized by the **source-generated** `Office365OnNewEmailTriggerPayload` from the Connectors SDK — no manual JSON parsing.
- `Body.Value` is a list — the gateway can batch multiple new emails per webhook call.
- Hover `payload.Body.Value[0]` to show the strongly-typed `GraphClientReceiveMessage`.

#### 🔴 BP 2 — The classifier decision
**[function-app/ImportanceClassifier.cs](../function-app/ImportanceClassifier.cs)** — top of `Classify()`, `var reasons = new List<string>();`

- Step through (**F10**) the three checks one by one and watch `reasons` grow.
- Hover `_importantSenders` — `HashSet<string>` populated once at startup from the `IMPORTANT_SENDERS` env var (case-insensitive lookup).
- At the `ScoreContent` call, step **into** (**F11**) — that's where the regex weights add up.

#### 🔴 BP 3 — Regex matching loop
**[function-app/ImportanceClassifier.cs](../function-app/ImportanceClassifier.cs)** — `foreach (var (pattern, weight, label) in UrgencyPatterns)`

- Watch `score` and `matchedLabels` grow. The first test payload (`"Urgent: action needed on connector demo"`) should hit "urgency adverb" and "explicit action/decision required".
- **Conditional breakpoint trick:** right-click → **Edit Breakpoint** → condition `pattern.IsMatch(haystack)` to stop only on hits.
- Talking point: adding a new signal is a one-line tuple in `UrgencyPatterns` — easy to extend.

#### 🔴 BP 4 — The skip path
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `if (!verdict.IsImportant)`

- After the urgent run, hit **Send Request** on the **second** POST in [test.http](../test.http) (newsletter). Verdict has zero reasons → 200 OK with no Teams call. Cheap negative path; keeps Teams quiet.

#### 🔴 BP 5 — Domain-based internal/external classification
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `response = await _graphClient.ListUsersAsync();`

- This is a **live call** to the Graph connector — proves managed-identity / `DefaultAzureCredential` is working locally.
- Step over and inspect `response.Value` — note ~100 users (the connector op doesn't expose paging). That's why we derive **tenant domains** instead of looking the user up directly.
- Hover `domains` after `ExtractDomains` — show the `HashSet` of verified domains.
- `isInternal` is `true` when the sender's domain is in the set — robust even when the user isn't in the first page.

#### 🔴 BP 6 — Best-effort enrichment
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var groupNames = await GetTopGroupNamesAsync(user.Id);`

- Step into. Two more Graph calls: `GetMemberGroupsAsync` then per-group `GetGroupPropertiesAsync`. Capped at `MaxGroupsToShow = 3` to bound latency and keep the card readable.
- Failures are caught + logged at Debug — the function never throws on enrichment errors.

#### 🔴 BP 7 — Teams card composition
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var request = new PostMessageRequest { ... }`

- Show the `PostMessageRequest : DynamicPostMessageRequest` derived class. The SDK uses an empty marker base + your subclass to serialize the operation's dynamic schema. **Zero hand-rolled JSON.**
- Inspect `messageBody` to show the assembled HTML with badge, role, groups, and "Why flagged".

#### 🔴 BP 8 — The actual outbound call
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var result = await _teamsClient.PostMessageToConversationAsync(...)`

- One line of code → an authenticated REST call to the Teams connection runtime URL via the gateway.
- Step over and switch to your Teams channel — the card appears in real time.
- Inspect `result.MessageID` — strongly-typed response.

### Bonus: catch connector failures fast

**Debug → Windows → Exception Settings** → tick `MsgraphgroupsanduserConnectorException` and `TeamsConnectorException`. If a connection isn't authorized, you'll stop right at the failure with `ex.StatusCode` visible.

### Bonus talking points to weave in

- **DI in [Program.cs](../function-app/Program.cs)** — `TeamsClient`, `MsgraphgroupsanduserClient`, and `ImportanceClassifier` are all singletons; the connector clients take `(connectionRuntimeUrl, managedIdentityClientId)`.
- **No `Microsoft.Graph` SDK** — we're using the `msgraphgroupsanduser` connector op set, not the official Graph SDK. The trade-off (no `/me/manager`, no paging) drove the domain-based design.
- **Deterministic local repro** — the typed trigger payload means [test.http](../test.http) is enough to exercise the full pipeline; no real Outlook needed during development.

---

## Part 2 — Cloud walkthrough

Now switch over to the deployed app to show how the same code runs in Azure and gets called by a real Office 365 email.

### Where things live

```
Subscription
└── rg-importantemailprocessor
    ├── func-<token>                  Function App (Flex Consumption)
    ├── plan-<token>                  Function App Plan (FC1)
    ├── id-<token>                    User-Assigned Managed Identity
    ├── st<token>                     Storage account (deployment + AzureWebJobsStorage)
    ├── log-<token>                   Log Analytics workspace
    ├── appi-<token>                  Application Insights
    └── cgw-<token>                   Connector Gateway  (lives in brazilsouth)
        ├── connections/
        │   ├── cgwc-<token>            Office 365 Outlook
        │   ├── cgwc-teams-<token>      Microsoft Teams
        │   └── cgwc-graph-<token>      Microsoft Graph (Groups & Users)
        └── triggerconfigs/
            └── cgwc-<token>-trigger    Office 365 OnNewEmailV3 → callback URL
```

Run `azd env get-values` to grab the live names if you need them.

### A) Tour of the Function App

[Azure Portal → Resource Group → Function App].

1. **Overview** — point out it's **Flex Consumption** (`FC1`, 2 GB memory, scale-to-zero) and the system-assigned + user-assigned MI.
2. **Settings → Environment variables** — show the relevant ones:
   - `TEAMS_TEAM_ID`, `TEAMS_CHANNEL_ID`, `IMPORTANT_SENDERS` — same values you saw in `local.settings.json`.
   - `TEAMS_CONNECTION_RUNTIME_URL`, `GRAPH_CONNECTION_RUNTIME_URL` — the gateway endpoints the SDK clients hit.
   - `AZURE_CLIENT_ID` — the UAMI client id; the connector clients use it to acquire tokens.
   - `APPLICATIONINSIGHTS_AUTHENTICATION_STRING: ClientId=...;Authorization=AAD` — AAD-only telemetry, no instrumentation key in plaintext.
3. **Settings → Identity → User assigned** — show the same UAMI; click into it to show its role assignments (Storage Blob Data Owner, Queue/Table Data Contributor on the storage account, Monitoring Metrics Publisher on App Insights).
4. **Functions → OnNewImportantEmailReceived → Code + Test → Logs** — leave it streaming for the real-email step below.
5. **Functions → App Keys → System keys** — show `connector_extension`. The post-deploy script grabbed this value to build the trigger callback URL — that's how the gateway is allowed to invoke the function.

### B) Tour of the Connector Gateway

[Azure Portal → Resource Group → cgw-`<token>`].

1. **Overview** — show it's a brand-new resource type (`Microsoft.Web/connectorGateways@2026-05-01-preview`) and that it lives in `brazilsouth` while the function lives in `westus2`. Cross-region by design today.
2. **Connections** — three rows:
   - `cgwc-<token>` (Office 365)
   - `cgwc-teams-<token>` (Teams)
   - `cgwc-graph-<token>` (Microsoft Graph — Groups & Users)

   Click each → **Status: Connected** (you authorized them after `azd up`). Show the **Access policies** tab — the Function App's UAMI principal id is allowed to use the connection. That's how the SDK clients get to call the runtime URL with no shared secret.
3. **Trigger configurations** → `cgwc-<token>-trigger`:
   - **Operation:** `OnNewEmailV3`
   - **Parameters:** `folderPath = Inbox` (note: no `importance` filter — we evaluate every mail in code via the classifier so we can use richer signals).
   - **Notification details / callback URL:** `https://func-<token>.azurewebsites.net/runtime/webhooks/connector?functionName=OnNewImportantEmailReceived&code=<system key>`. This URL was assembled by [infra/scripts/postdeploy.sh](../infra/scripts/postdeploy.sh) using the system key from step A.5.

### C) Tour of Application Insights

[Azure Portal → Resource Group → appi-`<token>`].

1. **Live metrics** — leave it open in a side panel for the real-email demo.
2. **Logs (KQL)** — handy queries to have ready:

   ```kusto
   // last 1h of important emails accepted
   traces
   | where timestamp > ago(1h)
   | where message has "Important email accepted"
   | project timestamp, message, severityLevel
   | order by timestamp desc
   ```

   ```kusto
   // skip-vs-accept ratio
   traces
   | where timestamp > ago(24h)
   | where message has_any ("Important email accepted", "Skipping non-important email")
   | summarize count() by tostring(split(message, ".")[0])
   ```

   ```kusto
   // any connector failures
   exceptions
   | where timestamp > ago(24h)
   | where type endswith "ConnectorException"
   | project timestamp, type, outerMessage, customDimensions
   ```

### D) Trigger a real run

1. **Pick a sender from the allowlist** (or temporarily add yourself) — confirm with `az functionapp config appsettings list -g rg-importantemailprocessor -n func-<token> --query "[?name=='IMPORTANT_SENDERS'].value" -o tsv`.
2. **Send yourself an email** from Outlook (web or desktop) to the inbox the Office 365 connection was authorized for. Two good test cases:
   - **Allowlist hit:** subject `Quick check on connector demo` from one of the `IMPORTANT_SENDERS`. Will trigger reason "Sender in IMPORTANT_SENDERS allowlist".
   - **Content heuristic hit:** any sender, subject `[URGENT] please review by EOD` — triggers "bracketed urgency tag" + "deadline language".
   - **Skip case:** any non-allowlisted sender, plain subject like `lunch tomorrow?` — should NOT post to Teams.
3. **What to watch happen, in order:**
   1. **Application Insights → Live metrics**: a new Request appears within ~5–30 s of the email landing (Office 365 trigger polling cadence).
   2. **Function App → Functions → OnNewImportantEmailReceived → Logs**: you'll see
      - `Important email accepted. Subject=... From=... Reasons=...` (for hits), or
      - `Skipping non-important email. Subject=... From=... Importance=...` (for misses).
   3. **Teams channel**: the triage card pops in with the badge (🟢 INTERNAL / 🔴 EXTERNAL), sender role + groups, "Why flagged", subject, and preview.
4. **Show the end-to-end correlation**: in App Insights → Transaction search, click the request → see the dependency calls out to `*.logic-df.azure-apihub.net` (Graph + Teams connection runtime URLs), with timing.

### E) Talking points for the cloud part

- **No secrets anywhere.** Function → connections is via UAMI access policy; function → storage / App Insights via UAMI role assignments; connection → external service via the user who authorized it in the gateway. Nothing in app settings is sensitive.
- **Server-side filter is just `folderPath = Inbox`.** All the importance logic is in your code — easy to evolve without redeploying gateway config.
- **Same code, two execution modes.** The local debug session and the cloud function run **identical** code paths and call the **same** connection runtime URLs. Local just uses your user identity instead of the UAMI.
- **Diagnostics surface naturally.** Every accept/skip is a structured log line, and connector failures throw typed exceptions you can pivot on in App Insights.

### F) Reset / cleanup

- Remove a sender from the allowlist:
  ```bash
  azd env set IMPORTANT_SENDERS "<comma list without that address>"
  azd deploy   # or azd provision if you only want to update settings
  ```
- Tear everything down:
  ```bash
  azd down --purge
  ```
