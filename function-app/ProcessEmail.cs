using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.Functions.Worker.Extensions.Connector;
using Microsoft.Azure.Connectors.DirectClient.Office365;
using Microsoft.Azure.Connectors.DirectClient.Teams;

namespace Company.Function
{
    public class ProcessEmail
    {
        private const string TeamsPostMessagePath = "/beta/teams/conversation/message/poster/Flow%20bot/location/Channel";

        // TODO: Replace with Azure OpenAI classification
        private static readonly HashSet<string> ImportantSenders = new(StringComparer.OrdinalIgnoreCase)
        {
            "thalme@microsoft.com",
            "paulyuk@microsoft.com",
            "nimashkowski@microsoft.com"
        };

        private static readonly string[] ImportantKeywords = ["urgent", "critical", "immediate action", "escalation"];

        private readonly ILogger _logger;
        private readonly TeamsClient _teamsClient;
        private readonly string _teamsTeamId;
        private readonly string _teamsChannelId;

        public ProcessEmail(ILoggerFactory loggerFactory, TeamsClient teamsClient)
        {
            _logger = loggerFactory.CreateLogger<ProcessEmail>();
            _teamsClient = teamsClient;
            _teamsTeamId = Environment.GetEnvironmentVariable("TEAMS_TEAM_ID") ?? "";
            _teamsChannelId = Environment.GetEnvironmentVariable("TEAMS_CHANNEL_ID") ?? "";
        }

        [Function("OnNewImportantEmailReceived")]
        public async Task<IActionResult> OnNewImportantEmailReceived(
            [ConnectorTrigger()] Office365OnNewEmailV3TriggerPayload payload)
        {
            var emails = payload.Body?.Value ?? [];

            foreach (var email in emails)
            {
                _logger.LogInformation("Subject: {Subject}, From: {From}, To: {To}",
                    email.Subject, email.From, email.To);

                if (!IsImportantEmail(email))
                {
                    _logger.LogInformation("Email not classified as important, skipping Teams notification.");
                    continue;
                }

                _logger.LogInformation("Email classified as important! Sending Teams notification.");

                await SendTeamsNotificationAsync(email);
            }

            return new OkResult();
        }

        /// <summary>
        /// Stub importance classifier. TODO: Replace with Azure OpenAI call.
        /// </summary>
        private static bool IsImportantEmail(GraphClientReceiveMessage email)
        {
            if (ImportantSenders.Contains(email.From ?? ""))
                return true;

            var subject = email.Subject ?? "";
            var body = email.Body ?? "";

            foreach (var keyword in ImportantKeywords)
            {
                if (subject.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                    body.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            if (string.Equals(email.Importance, "high", StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private async Task SendTeamsNotificationAsync(GraphClientReceiveMessage email)
        {
            if (string.IsNullOrEmpty(_teamsTeamId) || string.IsNullOrEmpty(_teamsChannelId))
            {
                _logger.LogWarning("TEAMS_TEAM_ID or TEAMS_CHANNEL_ID not configured. Skipping Teams notification.");
                return;
            }

            var messageBody = $"<b>📧 Important Email Received</b><br/>" +
                $"<b>From:</b> {email.From}<br/>" +
                $"<b>Subject:</b> {email.Subject}<br/>" +
                $"<b>Preview:</b> {email.BodyPreview ?? "(no preview)"}";

            var messagePayload = new
            {
                recipient = new
                {
                    groupId = _teamsTeamId,
                    channelId = _teamsChannelId
                },
                messageBody
            };
            var messageJson = JsonSerializer.Serialize(messagePayload);

            try
            {
                var result = await _teamsClient.SendRawRequestAsync<PostToConversationResponse>(
                    HttpMethod.Post,
                    TeamsPostMessagePath,
                    messageJson);

                _logger.LogInformation("Teams message posted. MessageId: {MessageId}", result?.MessageID);
            }
            catch (TeamsConnectorException ex)
            {
                _logger.LogError(ex, "Failed to post Teams message. Status: {StatusCode}", ex.StatusCode);
            }
        }
    }
}