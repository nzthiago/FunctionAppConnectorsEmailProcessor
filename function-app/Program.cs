using Azure.Core;
using Azure.Identity;
using Company.Function;
using Microsoft.Azure.Connectors.DirectClient.Msgraphgroupsanduser;
using Microsoft.Azure.Connectors.DirectClient.Teams;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureServices(services =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // DefaultAzureCredential works in both environments:
        //   - In Azure: uses the user-assigned managed identity whose client id is in AZURE_CLIENT_ID.
        //   - Locally:  falls back to the developer's `az login` / VS / VS Code credentials.
        // (Plain ManagedIdentityCredential would try IMDS at 169.254.169.254 — fine in Azure, fails locally.)
        var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
        {
            ManagedIdentityClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID")
        });

        services.AddSingleton<TokenCredential>(credential);

        services.AddSingleton(sp => new TeamsClient(
            Environment.GetEnvironmentVariable("TEAMS_CONNECTION_RUNTIME_URL") ?? "",
            sp.GetRequiredService<TokenCredential>(),
            httpClient: null));

        services.AddSingleton(sp => new MsgraphgroupsanduserClient(
            Environment.GetEnvironmentVariable("GRAPH_CONNECTION_RUNTIME_URL") ?? "",
            sp.GetRequiredService<TokenCredential>(),
            httpClient: null));

        services.AddSingleton<ImportanceClassifier>();
    })
    .Build();

host.Run();

