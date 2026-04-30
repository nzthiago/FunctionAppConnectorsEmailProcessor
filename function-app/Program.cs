using Azure.Core;
using Azure.Identity;
using Azure.Monitor.OpenTelemetry.Exporter;
using Company.Function;
using Microsoft.Azure.Connectors.DirectClient.Office365;
using Microsoft.Azure.Connectors.DirectClient.Teams;
using Microsoft.Azure.Functions.Worker.OpenTelemetry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;


// DefaultAzureCredential works in both environments:
//   - In Azure: uses the user-assigned managed identity whose client id is in AZURE_CLIENT_ID.
//   - Locally:  falls back to the developer's `az login` / VS / VS Code credentials.
// (Plain ManagedIdentityCredential would try IMDS at 169.254.169.254 — fine in Azure, fails locally.)
var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
{
    ManagedIdentityClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID")
});

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureServices(services =>
    {
        // OpenTelemetry → Azure Monitor (App Insights) for the worker process.
        // Paired with `"telemetryMode": "OpenTelemetry"` in host.json so host + worker
        // telemetry stays correlated. 
        services.AddOpenTelemetry()
            .UseFunctionsWorkerDefaults()
            .UseAzureMonitorExporter(o => o.Credential = credential);

        services.AddSingleton<TokenCredential>(credential);

        services.AddSingleton(sp => new TeamsClient(
            Environment.GetEnvironmentVariable("TEAMS_CONNECTION_RUNTIME_URL") ?? "",
            sp.GetRequiredService<TokenCredential>(),
            httpClient: null));

        // Office 365 client — used both for sender-history enrichment (GetEmailsAsync)
        // and to flag the source email (FlagAsync) once we decide it's important.
        // Same connection runtime URL the trigger uses; just consumed as a client too.
        services.AddSingleton(sp => new Office365Client(
            Environment.GetEnvironmentVariable("OFFICE365_CONNECTION_RUNTIME_URL") ?? "",
            sp.GetRequiredService<TokenCredential>(),
            httpClient: null));

        services.AddSingleton<ImportanceClassifier>();
    })
    .Build();

host.Run();

