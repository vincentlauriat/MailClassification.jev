using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using JevOutlook.Mail;

namespace JevOutlook.Graph;

public sealed record WellKnownFolders(string Inbox, string? Archive, string? JunkEmail, string? DeletedItems, string? Drafts, string? SentItems)
{
    public bool IsExcluded(string parentFolderId) =>
        parentFolderId == JunkEmail || parentFolderId == DeletedItems || parentFolderId == Drafts;
}

/// <summary>One inner response of a Graph JSON batch, mapped back to the request it answers.</summary>
public sealed class BatchPart
{
    public bool Ok { get; init; }
    public int Status { get; init; }
    public bool Retryable { get; init; }
    public bool AuthFailure { get; init; }
    public int RetryAfterMs { get; init; }
    public JsonElement Body { get; init; }
    public string Error { get; init; } = string.Empty;
}

public sealed record BatchRequest(string Id, string Method, string Url, IReadOnlyDictionary<string, string>? Headers = null, JsonNode? Body = null);

/// <summary>
/// Microsoft Graph implementation of <see cref="IMailbox"/> over <see cref="HttpClient"/>:
/// cursor-based listing, batched reads, master-category management, category
/// writes and archive moves. Labels are Outlook categories.
/// </summary>
public sealed class GraphMailClient : IMailbox
{
    public const string MetadataSelect =
        "id,conversationId,subject,from,toRecipients,ccRecipients,replyTo,receivedDateTime,bodyPreview,categories,hasAttachments,importance,inferenceClassification,internetMessageHeaders";
    public const string FullSelect = "id,body";
    private const string ListSelect = "id,receivedDateTime,categories,parentFolderId,isRead";
    private const string SortKeyFormat = "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'";

    private static readonly SemaphoreSlim MailboxGate = new(AppConstants.GraphMaxConcurrentCalls, AppConstants.GraphMaxConcurrentCalls);

    private readonly HttpClient _http;
    private readonly GraphTokenProvider _tokens;
    private WellKnownFolders? _folders;
    private MailboxCapabilities _capabilities = new("Outlook", "category", "categories", true, true,
        "Archive-eligible messages are moved to the Archive folder.");

    public GraphMailClient(HttpClient http, GraphTokenProvider tokens)
    {
        _http = http;
        _tokens = tokens;
    }

    public MailboxCapabilities Capabilities => _capabilities;

    /// <summary>Nothing to release: the HttpClient is shared and owned by the caller, tokens live in the OS cache.</summary>
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;

    // ----- Identity ---------------------------------------------------------

    public async Task<string> GetIdentityAsync(CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Get, "/me?$select=displayName,mail,userPrincipalName", null, null, ct);
        var json = await ReadJsonAsync(response, ct);
        var mail = MessageMetadata.GetString(json, "mail");
        if (mail.Length == 0) mail = MessageMetadata.GetString(json, "userPrincipalName");
        var name = MessageMetadata.GetString(json, "displayName");
        return name.Length > 0 ? $"{name} <{mail}>" : mail;
    }

    // ----- Folders ----------------------------------------------------------

    public async Task<MailboxCapabilities> ConnectAsync(CancellationToken ct)
    {
        var folders = await GetWellKnownFoldersAsync(ct);
        _capabilities = _capabilities with { CanArchive = folders.Archive is not null };
        return _capabilities;
    }

    public async Task<WellKnownFolders> GetWellKnownFoldersAsync(CancellationToken ct)
    {
        if (_folders is not null) return _folders;
        var requests = new List<BatchRequest>();
        foreach (var name in new[] { "inbox", "archive", "junkemail", "deleteditems", "drafts", "sentitems" })
        {
            requests.Add(new BatchRequest(name, "GET", $"/me/mailFolders/{name}?$select=id"));
        }
        var parts = await SendBatchAsync(requests, ct);
        string? IdOf(int index) => parts[index].Ok ? MessageMetadata.GetString(parts[index].Body, "id") : null;
        var inbox = IdOf(0) ?? throw new GraphRequestException((HttpStatusCode)parts[0].Status, "The Inbox folder could not be resolved: " + parts[0].Error);
        _folders = new WellKnownFolders(inbox, IdOf(1), IdOf(2), IdOf(3), IdOf(4), IdOf(5));
        return _folders;
    }

    /// <summary>Exact counts for the Inbox; approximate ($count) for the whole mailbox; -1 when unknown.</summary>
    public async Task<long> EstimateAsync(string scope, bool unreadOnly, CancellationToken ct)
    {
        try
        {
            if (scope == "inbox")
            {
                using var response = await SendAsync(HttpMethod.Get, "/me/mailFolders/inbox?$select=totalItemCount,unreadItemCount", null, null, ct);
                var json = await ReadJsonAsync(response, ct);
                var property = unreadOnly ? "unreadItemCount" : "totalItemCount";
                return json.TryGetProperty(property, out var value) && value.TryGetInt64(out var count) ? count : -1;
            }

            var url = "/me/messages?$count=true&$top=1&$select=id" + (unreadOnly ? "&$filter=" + Uri.EscapeDataString("isRead eq false") : string.Empty);
            using var all = await SendAsync(HttpMethod.Get, url, null, null, ct);
            var body = await ReadJsonAsync(all, ct);
            return body.TryGetProperty("@odata.count", out var odataCount) && odataCount.TryGetInt64(out var total) ? total : -1;
        }
        catch (GraphRequestException)
        {
            return -1;
        }
    }

    // ----- Listing ------------------------------------------------------------

    public Task<string> GetInitialCursorAsync(string scope, CancellationToken ct) =>
        Task.FromResult(FormatSortKey(DateTimeOffset.UtcNow.AddMinutes(5)));

    public static string FormatSortKey(DateTimeOffset when) => when.ToUniversalTime().ToString(SortKeyFormat, CultureInfo.InvariantCulture);

    public static DateTimeOffset ParseSortKey(string key) =>
        DateTimeOffset.TryParse(key, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var when)
            ? when
            : throw new MailboxException(false, $"The listing cursor '{key}' is not a valid timestamp. Clear the session and start again.");

    /// <summary>
    /// Newest-first page bounded by a receivedDateTime cursor. The caller keeps
    /// the ids seen at the boundary timestamp so ties are never lost or repeated.
    /// </summary>
    public async Task<ListPage> ListMessagesAsync(string scope, bool unreadOnly, string cursor, bool exclusive, int top, CancellationToken ct)
    {
        var folders = await GetWellKnownFoldersAsync(ct);
        var path = scope == "inbox" ? "/me/mailFolders/inbox/messages" : "/me/messages";
        var filter = BuildFilter(unreadOnly, ParseSortKey(cursor), exclusive);
        var url = $"{path}?$select={ListSelect}&$orderby={Uri.EscapeDataString("receivedDateTime desc")}&$top={top}&$filter={Uri.EscapeDataString(filter)}";
        using var response = await SendAsync(HttpMethod.Get, url, null, null, ct);
        var json = await ReadJsonAsync(response, ct);
        var items = new List<MessageRef>();
        if (json.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
        {
            foreach (var message in value.EnumerateArray())
            {
                var id = MessageMetadata.GetString(message, "id");
                var received = MessageMetadata.GetString(message, "receivedDateTime");
                if (id.Length == 0 || !DateTimeOffset.TryParse(received, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var when)) continue;
                items.Add(new MessageRef(id, FormatSortKey(when), MessageMetadata.ReadCategories(message),
                    folders.IsExcluded(MessageMetadata.GetString(message, "parentFolderId")),
                    message.TryGetProperty("isRead", out var isRead) && isRead.ValueKind == JsonValueKind.True));
            }
        }
        return new ListPage(items, items.Count >= top);
    }

    public static string BuildFilter(bool unreadOnly, DateTimeOffset cursor, bool exclusive = false)
    {
        var stamp = FormatSortKey(cursor);
        var filter = $"receivedDateTime {(exclusive ? "lt" : "le")} {stamp}";
        if (unreadOnly) filter += " and isRead eq false";
        return filter;
    }

    // ----- Batched reads --------------------------------------------------------

    /// <summary>Read many messages through Graph JSON batching; result order matches <paramref name="ids"/>.</summary>
    public async Task<List<ReadPart>> ReadMessagesAsync(IReadOnlyList<string> ids, ReadMode mode, CancellationToken ct)
    {
        var select = mode switch { ReadMode.Metadata => MetadataSelect, ReadMode.Full => FullSelect, _ => "id,categories" };
        var headers = mode == ReadMode.Full
            ? new Dictionary<string, string> { ["Prefer"] = "outlook.body-content-type=\"text\"" }
            : null;
        var requests = ids.Select((id, index) =>
            new BatchRequest(index.ToString(CultureInfo.InvariantCulture), "GET", $"/me/messages/{Uri.EscapeDataString(id)}?$select={select}", headers)).ToList();
        var parts = await SendBatchesAsync(requests, ct);
        return parts.Select(part => ToReadPart(part, mode)).ToList();
    }

    private static ReadPart ToReadPart(BatchPart part, ReadMode mode)
    {
        if (!part.Ok)
        {
            if (part.Status == 404) return ReadPart.Missing();
            return new ReadPart { Ok = false, Retryable = part.Retryable, AuthFailure = part.AuthFailure, RetryAfterMs = part.RetryAfterMs, Error = part.Error.Length > 0 ? part.Error : $"HTTP {part.Status}" };
        }
        switch (mode)
        {
            case ReadMode.Metadata:
            {
                var metadata = MessageMetadata.FromGraph(part.Body);
                return new ReadPart { Ok = true, Metadata = metadata, Labels = metadata.Categories };
            }
            case ReadMode.Full:
            {
                var (text, reason) = MessageContent.Extract(part.Body);
                return new ReadPart { Ok = true, FullText = text, FullReason = reason };
            }
            default:
                return new ReadPart { Ok = true, Labels = MessageMetadata.ReadCategories(part.Body) };
        }
    }

    // ----- Categories -------------------------------------------------------------

    /// <summary>
    /// Make sure every category exists in the master list (so it has a colour in
    /// Outlook). Returns the stored display names keyed by lower-case name.
    /// </summary>
    public async Task<Dictionary<string, string>> EnsureLabelsAsync(IReadOnlyList<string> names, CancellationToken ct)
    {
        var existing = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var next = "/me/outlook/masterCategories";
        for (var pages = 0; next is not null && pages < 50; pages++)
        {
            using var response = await SendAsync(HttpMethod.Get, next, null, null, ct);
            var json = await ReadJsonAsync(response, ct);
            if (json.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
            {
                foreach (var category in value.EnumerateArray())
                {
                    var name = MessageMetadata.GetString(category, "displayName");
                    if (name.Length > 0) existing.TryAdd(name, name);
                }
            }
            var link = MessageMetadata.GetString(json, "@odata.nextLink");
            next = link.StartsWith(AppConstants.GraphBaseUrl, StringComparison.OrdinalIgnoreCase) ? link[AppConstants.GraphBaseUrl.Length..] : null;
        }

        for (var i = 0; i < names.Count; i++)
        {
            var name = names[i];
            if (existing.ContainsKey(name)) continue;
            var color = AppConstants.CategoryColorPresets[i % AppConstants.CategoryColorPresets.Length];
            var body = new JsonObject { ["displayName"] = name, ["color"] = color };
            try
            {
                using var created = await SendAsync(HttpMethod.Post, "/me/outlook/masterCategories", body, null, ct);
                existing[name] = name;
            }
            catch (GraphRequestException ex) when (ex.Status == HttpStatusCode.Conflict)
            {
                existing[name] = name; // created concurrently
            }
        }

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in names) result[name.ToLowerInvariant()] = existing[name];
        return result;
    }

    /// <summary>Outlook categories are stored under the rule name itself.</summary>
    public string StoredLabel(string ruleName) => ruleName;

    // ----- Category cleanup -------------------------------------------------------------

    /// <summary>Ids (+ current categories) of up to <paramref name="top"/> messages carrying <paramref name="label"/>, anywhere in the mailbox.</summary>
    public async Task<List<(string Id, IReadOnlyList<string> Labels)>> FindMessagesWithLabelAsync(string label, int top, CancellationToken ct)
    {
        var filter = $"categories/any(c:c eq '{label.Replace("'", "''")}')";
        var url = $"/me/messages?$select=id,categories&$top={top}&$filter={Uri.EscapeDataString(filter)}";
        using var response = await SendAsync(HttpMethod.Get, url, null, null, ct);
        var json = await ReadJsonAsync(response, ct);
        var items = new List<(string, IReadOnlyList<string>)>();
        if (json.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
        {
            foreach (var message in value.EnumerateArray())
            {
                var id = MessageMetadata.GetString(message, "id");
                if (id.Length > 0) items.Add((id, MessageMetadata.ReadCategories(message)));
            }
        }
        return items;
    }

    /// <summary>Delete a category from the master list (messages keep the name until they are patched).</summary>
    public async Task<bool> DeleteLabelAsync(string label, CancellationToken ct)
    {
        var next = "/me/outlook/masterCategories";
        for (var pages = 0; next is not null && pages < 50; pages++)
        {
            using var response = await SendAsync(HttpMethod.Get, next, null, null, ct);
            var json = await ReadJsonAsync(response, ct);
            if (json.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
            {
                foreach (var entry in value.EnumerateArray())
                {
                    if (!string.Equals(MessageMetadata.GetString(entry, "displayName"), label, StringComparison.OrdinalIgnoreCase)) continue;
                    var id = MessageMetadata.GetString(entry, "id");
                    using var deleted = await SendAsync(HttpMethod.Delete, $"/me/outlook/masterCategories/{Uri.EscapeDataString(id)}", null, null, ct);
                    return true;
                }
            }
            var link = MessageMetadata.GetString(json, "@odata.nextLink");
            next = link.StartsWith(AppConstants.GraphBaseUrl, StringComparison.OrdinalIgnoreCase) ? link[AppConstants.GraphBaseUrl.Length..] : null;
        }
        return false;
    }

    // ----- Writes ---------------------------------------------------------------------

    /// <summary>Replace the category list of each message (idempotent; a vanished message counts as done).</summary>
    public Task<WriteOutcome> ApplyLabelsAsync(IReadOnlyList<(string Id, IReadOnlyList<string> Labels)> updates, CancellationToken ct)
    {
        var requests = updates.Select((u, index) => new BatchRequest(
            index.ToString(CultureInfo.InvariantCulture), "PATCH", $"/me/messages/{Uri.EscapeDataString(u.Id)}",
            new Dictionary<string, string> { ["Content-Type"] = "application/json" },
            new JsonObject { ["categories"] = new JsonArray(u.Labels.Select(c => (JsonNode)JsonValue.Create(c)!).ToArray()) })).ToList();
        return ExecuteWritesAsync(requests, "update categories", ct);
    }

    /// <summary>Remove one category: PATCH each message with its categories minus <paramref name="label"/>.</summary>
    public Task<WriteOutcome> RemoveLabelAsync(IReadOnlyList<(string Id, IReadOnlyList<string> Labels)> messages, string label, CancellationToken ct) =>
        ApplyLabelsAsync(messages.Select(m => (m.Id, (IReadOnlyList<string>)m.Labels
            .Where(c => !string.Equals(c, label, StringComparison.OrdinalIgnoreCase)).ToList())).ToList(), ct);

    /// <summary>Move messages to the well-known Archive folder (the Outlook equivalent of removing INBOX).</summary>
    public Task<WriteOutcome> ArchiveAsync(IReadOnlyList<string> ids, CancellationToken ct)
    {
        var requests = ids.Select((id, index) => new BatchRequest(
            index.ToString(CultureInfo.InvariantCulture), "POST", $"/me/messages/{Uri.EscapeDataString(id)}/move",
            new Dictionary<string, string> { ["Content-Type"] = "application/json" },
            new JsonObject { ["destinationId"] = "archive" })).ToList();
        return ExecuteWritesAsync(requests, "archive", ct);
    }

    /// <summary>
    /// Grouped idempotent writes. Exchange throttles writes far harder than reads, so
    /// batches go out one at a time and throttled requests are retried with the
    /// server's Retry-After (or an exponential fallback) up to <see cref="AppConstants.GraphWriteRetries"/> times.
    /// </summary>
    private async Task<WriteOutcome> ExecuteWritesAsync(List<BatchRequest> requests, string action, CancellationToken ct)
    {
        if (requests.Count == 0) return WriteOutcome.Success;
        var pending = requests;
        for (var attempt = 0; ; attempt++)
        {
            var parts = new List<BatchPart>();
            foreach (var chunk in pending.Chunk(AppConstants.GraphBatchRequestLimit))
            {
                parts.AddRange(await SendBatchAsync(chunk, ct));
            }
            var retry = new List<BatchRequest>();
            var retryAfter = 0;
            var lastStatus = 0;
            for (var i = 0; i < pending.Count; i++)
            {
                var part = parts[i];
                if (part.Ok || part.Status == 404) continue; // 404: message moved/deleted meanwhile — nothing left to do
                if (part.AuthFailure)
                    return new WriteOutcome(false, false, "Microsoft Graph authorization failed while trying to " + action + " messages. Sign in again (jevoutlook account login).");
                if (!part.Retryable)
                    return new WriteOutcome(false, false, $"Microsoft Graph rejected a request to {action} a message with HTTP {part.Status}. {part.Error}".Trim());
                retryAfter = Math.Max(retryAfter, part.RetryAfterMs);
                lastStatus = part.Status;
                retry.Add(pending[i]);
            }
            if (retry.Count == 0) return WriteOutcome.Success;
            if (attempt >= AppConstants.GraphWriteRetries)
            {
                return new WriteOutcome(false, true,
                    $"Microsoft Graph is temporarily rate-limiting or unavailable ({retry.Count} {action} request(s) still failing with HTTP {lastStatus} after {attempt} retries).");
            }
            var fallback = AppConstants.GraphRetryDelayMs * (1 << attempt) * 2; // 1 s, 2 s, 4 s, 8 s, 16 s
            await Task.Delay(Math.Min(AppConstants.MaxGraphWriteRetryDelayMs, Math.Max(retryAfter, fallback)), ct);
            pending = retry;
        }
    }

    // ----- Batch plumbing -----------------------------------------------------------------

    /// <summary>Split into batches of 20 and run them concurrently (bounded by the mailbox gate).</summary>
    public async Task<List<BatchPart>> SendBatchesAsync(IReadOnlyList<BatchRequest> requests, CancellationToken ct)
    {
        var chunks = requests.Chunk(AppConstants.GraphBatchRequestLimit).ToList();
        var tasks = chunks.Select(chunk => SendBatchAsync(chunk, ct)).ToList();
        var results = await Task.WhenAll(tasks);
        return results.SelectMany(r => r).ToList();
    }

    /// <summary>One Graph $batch call. Responses are mapped by id so server-side reordering is harmless.</summary>
    public async Task<List<BatchPart>> SendBatchAsync(IReadOnlyList<BatchRequest> requests, CancellationToken ct)
    {
        var payload = new JsonObject
        {
            ["requests"] = new JsonArray(requests.Select(r =>
            {
                var node = new JsonObject { ["id"] = r.Id, ["method"] = r.Method, ["url"] = r.Url };
                if (r.Headers is { Count: > 0 })
                {
                    var headers = new JsonObject();
                    foreach (var (k, v) in r.Headers) headers[k] = v;
                    node["headers"] = headers;
                }
                if (r.Body is not null) node["body"] = r.Body.DeepClone();
                return (JsonNode)node;
            }).ToArray()),
        };

        HttpResponseMessage response;
        try
        {
            response = await SendAsync(HttpMethod.Post, "/$batch", payload, null, ct, throwOnError: false);
        }
        catch (Exception ex) when (ex is HttpRequestException || (ex is TaskCanceledException && !ct.IsCancellationRequested))
        {
            return requests.Select(_ => new BatchPart { Ok = false, Status = 0, Retryable = true, Error = "Microsoft Graph could not be reached: " + ex.Message }).ToList();
        }

        using (response)
        {
            var status = (int)response.StatusCode;
            if (status == 401 || status == 403)
            {
                return requests.Select(_ => new BatchPart { Ok = false, Status = status, AuthFailure = true, Error = "Authorization failed." }).ToList();
            }
            if (status < 200 || status >= 300)
            {
                var retryable = IsTransient(status);
                var retryAfter = RetryAfterMs(response.Headers);
                Debug($"$batch itself returned HTTP {status} (Retry-After {retryAfter} ms) for {requests.Count} request(s)");
                return requests.Select(_ => new BatchPart { Ok = false, Status = status, Retryable = retryable, RetryAfterMs = retryAfter, Error = $"Graph batch returned HTTP {status}." }).ToList();
            }

            var json = await ReadJsonAsync(response, ct);
            var byId = new Dictionary<string, BatchPart>(StringComparer.Ordinal);
            if (json.TryGetProperty("responses", out var responses) && responses.ValueKind == JsonValueKind.Array)
            {
                foreach (var inner in responses.EnumerateArray())
                {
                    var id = MessageMetadata.GetString(inner, "id");
                    var innerStatus = inner.TryGetProperty("status", out var s) && s.TryGetInt32(out var code) ? code : 0;
                    var ok = innerStatus >= 200 && innerStatus < 300;
                    var body = inner.TryGetProperty("body", out var b) ? b.Clone() : default;
                    var retryAfter = 0;
                    if (inner.TryGetProperty("headers", out var hs) && hs.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var h in hs.EnumerateObject())
                        {
                            if (h.Name.Equals("Retry-After", StringComparison.OrdinalIgnoreCase)) retryAfter = ParseRetryAfter(h.Value.GetString());
                        }
                    }
                    byId[id] = new BatchPart
                    {
                        Ok = ok,
                        Status = innerStatus,
                        Body = body,
                        Retryable = !ok && IsTransient(innerStatus),
                        AuthFailure = innerStatus == 401,
                        RetryAfterMs = retryAfter,
                        Error = ok ? string.Empty : ExtractGraphError(body, innerStatus),
                    };
                }
            }
            if (DebugEnabled)
            {
                var failed = byId.Values.Where(p => !p.Ok && p.Status != 404).ToList();
                if (failed.Count > 0)
                    Debug($"$batch of {requests.Count} ({requests[0].Method}): {failed.Count} failed — " +
                          string.Join("; ", failed.GroupBy(p => p.Status).Select(g => $"HTTP {g.Key} ×{g.Count()} Retry-After {g.Max(p => p.RetryAfterMs)} ms: {g.First().Error}")));
            }

            return requests.Select(r => byId.TryGetValue(r.Id, out var part)
                ? part
                : new BatchPart { Ok = false, Status = 0, Retryable = true, Error = "Microsoft Graph returned an incomplete batch response." }).ToList();
        }
    }

    // ----- HTTP helpers ------------------------------------------------------------------------

    private async Task<HttpResponseMessage> SendAsync(HttpMethod method, string relativeUrl, JsonNode? body,
        IReadOnlyDictionary<string, string>? headers, CancellationToken ct, bool throwOnError = true)
    {
        for (var attempt = 0; ; attempt++)
        {
            await MailboxGate.WaitAsync(ct);
            var released = false;
            HttpResponseMessage response;
            try
            {
                using var request = new HttpRequestMessage(method, AppConstants.GraphBaseUrl + relativeUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await _tokens.GetAccessTokenAsync(ct));
                if (headers is not null) foreach (var (k, v) in headers) request.Headers.TryAddWithoutValidation(k, v);
                if (body is not null) request.Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json");
                response = await _http.SendAsync(request, HttpCompletionOption.ResponseContentRead, ct);
            }
            catch (Exception ex) when (ex is HttpRequestException || (ex is TaskCanceledException && !ct.IsCancellationRequested))
            {
                if (attempt < AppConstants.NetworkRetries)
                {
                    MailboxGate.Release();
                    released = true;
                    await Task.Delay(AppConstants.GraphRetryDelayMs, ct);
                    continue;
                }
                throw new GraphRequestException(0, "Microsoft Graph could not be reached: " + ex.Message);
            }
            finally
            {
                if (!released) MailboxGate.Release();
            }

            var status = (int)response.StatusCode;
            if (throwOnError && attempt < AppConstants.GraphReadRetries && IsTransient(status))
            {
                Debug($"HTTP {status} on {method} {relativeUrl} (Retry-After {RetryAfterMs(response.Headers)} ms), attempt {attempt + 1}");
                var delay = Math.Min(AppConstants.MaxGraphRetryDelayMs, Math.Max(AppConstants.GraphRetryDelayMs * (1 << attempt) * 2, RetryAfterMs(response.Headers)));
                response.Dispose();
                await Task.Delay(delay, ct);
                continue;
            }
            if (throwOnError && (status < 200 || status >= 300))
            {
                var text = await response.Content.ReadAsStringAsync(ct);
                response.Dispose();
                JsonElement error = default;
                try { error = JsonDocument.Parse(text).RootElement.Clone(); } catch (JsonException) { }
                throw new GraphRequestException((HttpStatusCode)status, ExtractGraphError(error, status));
            }
            return response;
        }
    }

    private static async Task<JsonElement> ReadJsonAsync(HttpResponseMessage response, CancellationToken ct)
    {
        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct);
        return doc.RootElement.Clone();
    }

    public static bool IsTransient(int status) => status == 408 || status == 409 || status == 425 || status == 429 || status >= 500 || status == 0;

    private static readonly bool DebugEnabled = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("JEVOUTLOOK_DEBUG"));

    private static void Debug(string message)
    {
        if (DebugEnabled) Console.Error.WriteLine($"[graph {DateTime.Now:HH:mm:ss}] {message}");
    }

    private static int RetryAfterMs(HttpResponseHeaders headers)
    {
        if (headers.RetryAfter is null) return 0;
        if (headers.RetryAfter.Delta is { } delta) return (int)Math.Max(0, delta.TotalMilliseconds);
        if (headers.RetryAfter.Date is { } date) return (int)Math.Max(0, (date - DateTimeOffset.UtcNow).TotalMilliseconds);
        return 0;
    }

    public static int ParseRetryAfter(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return 0;
        if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var seconds) && seconds >= 0) return (int)(seconds * 1000);
        if (DateTimeOffset.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var date))
            return (int)Math.Max(0, (date - DateTimeOffset.UtcNow).TotalMilliseconds);
        return 0;
    }

    private static string ExtractGraphError(JsonElement body, int status)
    {
        if (body.ValueKind == JsonValueKind.Object && body.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.Object)
        {
            var code = MessageMetadata.GetString(error, "code");
            var message = MessageMetadata.GetString(error, "message");
            var text = (code + " " + message).Trim();
            if (text.Length > 0) return text.Length > 300 ? text[..300] : text;
        }
        return $"HTTP {status}";
    }
}

public sealed class GraphRequestException(HttpStatusCode status, string message)
    : MailboxException(GraphMailClient.IsTransient((int)status), message)
{
    public HttpStatusCode Status { get; } = status;
}
