using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.Functions.Worker.Extensions.Connector;
using Microsoft.Azure.Connectors.DirectClient.Office365;

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
        public async Task<IActionResult> OnNewImportantEmailReceived(
            [ConnectorTrigger()] Office365OnNewEmailV3TriggerPayload payload)
        {
            var emails = payload.Body?.Value ?? [];

            foreach (var email in emails)
            {
                _logger.LogInformation("Subject: {Subject}, From: {From}, To: {To}, Body: {Body}",
                    email.Subject, email.From, email.To, email.Body);
            }
            return new OkResult();
        }
    }
}