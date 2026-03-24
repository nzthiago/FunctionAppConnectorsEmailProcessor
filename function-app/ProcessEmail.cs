using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace Company.Function
{
    public class ProcessEmail
    {
        private readonly ILogger _logger;

        public ProcessEmail(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<ProcessEmail>();
        }

        [Function("OnNewImportantEmailReceived")]
        public async Task<IActionResult> OnNewImportantEmailReceived([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
        {
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            _logger.LogInformation("C# HTTP POST trigger function processed a request with body: {RequestBody}", requestBody);
            return new OkResult();
        }
    }
}