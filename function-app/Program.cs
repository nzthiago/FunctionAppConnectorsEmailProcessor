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

        services.AddSingleton(_ =>
        {
            var connectionRuntimeUrl = Environment.GetEnvironmentVariable("TEAMS_CONNECTION_RUNTIME_URL") ?? "";
            var managedIdentityClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
            return new TeamsClient(connectionRuntimeUrl, managedIdentityClientId);
        });

        services.AddSingleton(_ =>
        {
            var connectionRuntimeUrl = Environment.GetEnvironmentVariable("GRAPH_CONNECTION_RUNTIME_URL") ?? "";
            var managedIdentityClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
            return new MsgraphgroupsanduserClient(connectionRuntimeUrl, managedIdentityClientId);
        });

        services.AddSingleton<ImportanceClassifier>();
    })
    .Build();

host.Run();
