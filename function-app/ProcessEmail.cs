using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.Functions.Worker.Extensions.Connector;
using Microsoft.Azure.Connectors.DirectClient.Office365;
using Microsoft.Azure.Connectors.DirectClient.Msgraphgroupsanduser;
using Microsoft.Azure.Connectors.DirectClient.Teams;

namespace Company.Function
{
    /// <summary>
    /// Triage bot for inbound mail. The Office 365 trigger fires for every new message
    /// in the watched folder; this function decides — via <see cref="ImportanceClassifier"/>
    /// — whether each one is important enough to surface in the Teams triage channel.
    ///
    /// Important means ANY of:
    ///   * Sender is in the IMPORTANT_SENDERS allowlist
    ///   * Email is flagged High importance by the sender
    ///   * Subject/body contains urgency / action-required language above threshold
    ///
    /// For mails that pass the bar, we use Microsoft Graph (Groups & Users connector) to:
    ///   * Look up the sender in Microsoft Entra ID (best-effort enrichment)
    ///   * If internal, pull job title, department, and group memberships
    ///   * If external, badge the Teams notification with a phishing-aware warning
    /// The enriched payload is then posted to a Teams channel as a structured card.
    /// </summary>
    public class ProcessEmail
    {
        private const string PostAsFlowBot = "Flow bot";
        private const string PostInChannel = "Channel";
        private const int MaxGroupsToShow = 3;

        // Derived from the empty marker DynamicPostMessageRequest so the SDK serializes
        // the runtime fields below into the dynamic-schema POST body the connector expects.
        private sealed class PostMessageRequest : DynamicPostMessageRequest
        {
            public RecipientInfo Recipient { get; set; } = new();
            public string MessageBody { get; set; } = string.Empty;
        }

        private sealed class RecipientInfo
        {
            public string GroupId { get; set; } = string.Empty;
            public string ChannelId { get; set; } = string.Empty;
        }

        // Subset of the user fields returned by Graph ListUsers we care about.
        private sealed class GraphUser
        {
            [JsonPropertyName("id")] public string? Id { get; set; }
            [JsonPropertyName("displayName")] public string? DisplayName { get; set; }
            [JsonPropertyName("jobTitle")] public string? JobTitle { get; set; }
            [JsonPropertyName("department")] public string? Department { get; set; }
            [JsonPropertyName("mail")] public string? Mail { get; set; }
            [JsonPropertyName("userPrincipalName")] public string? UserPrincipalName { get; set; }
        }

        private sealed record SenderContext(
            bool IsInternal,
            string? DisplayName,
            string? JobTitle,
            string? Department,
            IReadOnlyList<string> GroupNames);

        private static readonly JsonSerializerOptions JsonOpts = new()
        {
            PropertyNameCaseInsensitive = true
        };

        private readonly ILogger _logger;
        private readonly TeamsClient _teamsClient;
        private readonly MsgraphgroupsanduserClient _graphClient;
        private readonly ImportanceClassifier _classifier;
        private readonly string _teamsTeamId;
        private readonly string _teamsChannelId;

        public ProcessEmail(
            ILoggerFactory loggerFactory,
            TeamsClient teamsClient,
            MsgraphgroupsanduserClient graphClient,
            ImportanceClassifier classifier)
        {
            _logger = loggerFactory.CreateLogger<ProcessEmail>();
            _teamsClient = teamsClient;
            _graphClient = graphClient;
            _classifier = classifier;
            _teamsTeamId = Environment.GetEnvironmentVariable("TEAMS_TEAM_ID") ?? "";
            _teamsChannelId = Environment.GetEnvironmentVariable("TEAMS_CHANNEL_ID") ?? "";
        }

        [Function("OnNewImportantEmailReceived")]
        public async Task<IActionResult> OnNewImportantEmailReceived(
            [ConnectorTrigger()] Office365OnNewEmailTriggerPayload payload)
        {
            var emails = payload.Body?.Value ?? [];

            foreach (var email in emails)
            {
                var verdict = _classifier.Classify(
                    email.From,
                    email.Subject,
                    email.Body,
                    email.BodyPreview,
                    email.Importance);

                if (!verdict.IsImportant)
                {
                    _logger.LogInformation(
                        "Skipping non-important email. Subject={Subject} From={From} Importance={Importance}",
                        email.Subject, email.From, email.Importance);
                    continue;
                }

                _logger.LogInformation(
                    "Important email accepted. Subject={Subject} From={From} Reasons={Reasons}",
                    email.Subject, email.From, string.Join(" | ", verdict.Reasons));

                var sender = await ResolveSenderContextAsync(email.From);
                await PostTriageCardAsync(email, sender, verdict);
            }

            return new OkResult();
        }

        /// <summary>
        /// Resolves the sender's directory context. Internal senders are augmented with
        /// org info and group memberships; external senders return IsInternal=false so
        /// the Teams card can flag them appropriately. Best-effort: failures degrade
        /// gracefully to "external" rather than throwing.
        /// </summary>
        private async Task<SenderContext> ResolveSenderContextAsync(string? senderEmail)
        {
            var unknown = new SenderContext(false, null, null, null, Array.Empty<string>());
            if (string.IsNullOrWhiteSpace(senderEmail))
            {
                return unknown;
            }

            var senderDomain = senderEmail.Contains('@')
                ? senderEmail[(senderEmail.IndexOf('@') + 1)..]
                : null;

            ListUsersResponse? response;
            try
            {
                // The Graph connector's ListUsers op hits /v1.0/users with no $filter,
                // $top, or $skiptoken — so for large tenants we only see the first page
                // (~100 users). That makes per-user lookup unreliable. Instead we use
                // the page to derive the tenant's verified domain set, and classify the
                // sender as INTERNAL when their domain is one of those. Per-user
                // enrichment (name/title/groups) is best-effort: only populated when the
                // sender happens to be in the returned page.
                response = await _graphClient.ListUsersAsync();
            }
            catch (MsgraphgroupsanduserConnectorException ex)
            {
                _logger.LogWarning(ex,
                    "Graph ListUsers failed (status {StatusCode}) — treating sender {Sender} as external.",
                    ex.StatusCode, senderEmail);
                return unknown;
            }

            var domains = ExtractDomains(response?.Value);
            var isInternal = senderDomain is not null && domains.Contains(senderDomain);
            if (!isInternal)
            {
                return unknown;
            }

            // Best-effort enrichment: only if sender is in the returned page.
            var user = FindUserByEmail(response?.Value, senderEmail);
            if (user is null || string.IsNullOrEmpty(user.Id))
            {
                return new SenderContext(true, null, null, null, Array.Empty<string>());
            }

            var groupNames = await GetTopGroupNamesAsync(user.Id);
            return new SenderContext(true, user.DisplayName, user.JobTitle, user.Department, groupNames);
        }

        private static HashSet<string> ExtractDomains(List<object>? users)
        {
            var domains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (users is null) return domains;

            foreach (var raw in users)
            {
                var u = TryDeserializeUser(raw);
                if (u is null) continue;

                var address = !string.IsNullOrWhiteSpace(u.Mail) ? u.Mail : u.UserPrincipalName;
                if (string.IsNullOrWhiteSpace(address)) continue;

                var atIdx = address.IndexOf('@');
                if (atIdx > 0 && atIdx < address.Length - 1)
                {
                    domains.Add(address[(atIdx + 1)..]);
                }
            }
            return domains;
        }

        private static GraphUser? FindUserByEmail(List<object>? users, string senderEmail)
        {
            if (users is null) return null;

            foreach (var raw in users)
            {
                var candidate = TryDeserializeUser(raw);
                if (candidate is null) continue;

                if (string.Equals(candidate.Mail, senderEmail, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(candidate.UserPrincipalName, senderEmail, StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }
            return null;
        }

        private static GraphUser? TryDeserializeUser(object raw)
        {
            try
            {
                var json = raw is JsonElement el ? el.GetRawText() : JsonSerializer.Serialize(raw);
                return JsonSerializer.Deserialize<GraphUser>(json, JsonOpts);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Hydrates the user's group membership IDs into display names. We cap at
        /// MaxGroupsToShow to keep the Teams card readable and the per-call latency
        /// bounded (one Graph call per group resolved).
        /// </summary>
        private async Task<IReadOnlyList<string>> GetTopGroupNamesAsync(string userObjectId)
        {
            List<string>? groupIds;
            try
            {
                var memberOf = await _graphClient.GetMemberGroupsAsync(
                    userObjectId,
                    new GetMemberGroupsInput { SecurityEnabledOnly = false });
                groupIds = memberOf?.Value;
            }
            catch (MsgraphgroupsanduserConnectorException ex)
            {
                _logger.LogWarning(ex, "Graph GetMemberGroups failed for user {UserId}.", userObjectId);
                return Array.Empty<string>();
            }

            if (groupIds is null || groupIds.Count == 0)
            {
                return Array.Empty<string>();
            }

            var names = new List<string>();
            foreach (var groupId in groupIds.Take(MaxGroupsToShow))
            {
                try
                {
                    var group = await _graphClient.GetGroupPropertiesAsync(groupId);
                    if (!string.IsNullOrWhiteSpace(group?.DisplayName))
                    {
                        names.Add(group.DisplayName);
                    }
                }
                catch (MsgraphgroupsanduserConnectorException ex)
                {
                    _logger.LogDebug(ex, "Could not resolve group {GroupId}.", groupId);
                }
            }
            return names;
        }

        private async Task PostTriageCardAsync(GraphClientReceiveMessage email, SenderContext sender, ImportanceVerdict verdict)
        {
            if (string.IsNullOrEmpty(_teamsTeamId) || string.IsNullOrEmpty(_teamsChannelId))
            {
                _logger.LogWarning("TEAMS_TEAM_ID or TEAMS_CHANNEL_ID not configured. Skipping Teams notification.");
                return;
            }

            var badge = sender.IsInternal
                ? "🟢 <b>INTERNAL</b>"
                : "🔴 <b>EXTERNAL — verify identity before acting</b>";

            var senderLine = sender.IsInternal && !string.IsNullOrWhiteSpace(sender.DisplayName)
                ? $"<b>From:</b> {sender.DisplayName} &lt;{email.From}&gt;"
                : $"<b>From:</b> {email.From}";

            var roleLine = sender.IsInternal && (!string.IsNullOrWhiteSpace(sender.JobTitle) || !string.IsNullOrWhiteSpace(sender.Department))
                ? $"<br/><b>Role:</b> {sender.JobTitle}{(string.IsNullOrWhiteSpace(sender.Department) ? "" : $" — {sender.Department}")}"
                : "";

            var groupsLine = sender.GroupNames.Count > 0
                ? $"<br/><b>Teams / groups:</b> {string.Join(", ", sender.GroupNames)}"
                : "";

            var reasonsLine = verdict.Reasons.Count > 0
                ? $"<br/><b>Why flagged:</b> {string.Join("; ", verdict.Reasons)}"
                : "";

            var messageBody =
                $"<b>📧 Email triage — review required</b><br/>" +
                $"{badge}<br/>" +
                $"{senderLine}{roleLine}{groupsLine}{reasonsLine}<br/>" +
                $"<b>Subject:</b> {email.Subject}<br/>" +
                $"<b>Preview:</b> {email.BodyPreview ?? "(no preview)"}";

            var request = new PostMessageRequest
            {
                Recipient = new RecipientInfo
                {
                    GroupId = _teamsTeamId,
                    ChannelId = _teamsChannelId
                },
                MessageBody = messageBody
            };

            try
            {
                var result = await _teamsClient.PostMessageToConversationAsync(
                    PostAsFlowBot,
                    PostInChannel,
                    request);

                _logger.LogInformation("Triage card posted to Teams. MessageId={MessageId}", result?.MessageID);
            }
            catch (TeamsConnectorException ex)
            {
                _logger.LogError(ex, "Failed to post Teams message. Status={StatusCode}", ex.StatusCode);
            }
        }
    }
}
