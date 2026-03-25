using Azure.Identity;
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
    })
    .Build();

host.Run();
