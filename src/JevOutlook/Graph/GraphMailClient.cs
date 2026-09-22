using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace JevOutlook.Graph;

public sealed record MessageRef(string Id, DateTimeOffset ReceivedDateTime, IReadOnlyList<string> Categories, string ParentFolderId, bool IsRead);

public sealed record ListPage(IReadOnlyList<MessageRef> Items, bool Full);

public sealed record WellKnownFolders(string Inbox, string? Archive, string? JunkEmail, string? DeletedItems, string? Drafts, string? SentItems)
{
    public bool IsExcluded(string parentFolderId) =>
        parentFolderId == JunkEmail || parentFolderId == DeletedItems || parentFolderId == Drafts;
}

public enum ReadMode { Metadata, Full, Categories }

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

public sealed record WriteOutcome(bool Ok, bool Retryable, string Error)
{
    public static readonly WriteOutcome Success = new(true, false, string.Empty);
}

public sealed record BatchRequest(string Id, string Method, string Url, IReadOnlyDictionary<string, string>? Headers = null, JsonNode? Body = null);

/// <summary>
/// Thin Microsoft Graph mail client over <see cref="HttpClient"/>. It exposes
/// exactly the operations the triage engine needs: cursor-based listing,
/// batched reads, master-category management, category writes and archive moves.
/// </summary>
public sealed class GraphMailClient
{
    public const string MetadataSelect =
        "id,conversationId,subject,from,toRecipients,ccRecipients,replyTo,receivedDateTime,bodyPreview,categories,hasAttachments,importance,inferenceClassification,internetMessageHeaders";
    public const string FullSelect = "id,body";
    private const string ListSelect = "id,receivedDateTime,categories,parentFolderId,isRead";

    private static readonly SemaphoreSlim MailboxGate = new(AppConstants.GraphMaxConcurrentCalls, AppConstants.GraphMaxConcurrentCalls);

    private readonly HttpClient _http;
    private readonly GraphTokenProvider _tokens;

    public GraphMailClient(HttpClient http, GraphTokenProvider tokens)
    {
        _http = http;
        _tokens = tokens;
    }

    // ----- Identity ---------------------------------------------------------

    public async Task<string> GetSignedInUserAsync(CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Get, "/me?$select=displayName,mail,userPrincipalName", null, null, ct);
        var json = await ReadJsonAsync(response, ct);
        var mail = MessageMetadata.GetString(json, "mail");
        if (mail.Length == 0) mail = MessageMetadata.GetString(json, "userPrincipalName");
        var name = MessageMetadata.GetString(json, "displayName");
        return name.Length > 0 ? $"{name} <{mail}>" : mail;
    }

    // ----- Folders ----------------------------------------------------------

    public async Task<WellKnownFolders> GetWellKnownFoldersAsync(CancellationToken ct)
    {
        var requests = new List<BatchRequest>();
        foreach (var name in new[] { "inbox", "archive", "junkemail", "deleteditems", "drafts", "sentitems" })
        {
            requests.Add(new BatchRequest(name, "GET", $"/me/mailFolders/{name}?$select=id"));
        }
        var parts = await SendBatchAsync(requests, ct);
        string? IdOf(int index) => parts[index].Ok ? MessageMetadata.GetString(parts[index].Body, "id") : null;
        var inbox = IdOf(0) ?? throw new InvalidOperationException("The Inbox folder could not be resolved: " + parts[0].Error);
        return new WellKnownFolders(inbox, IdOf(1), IdOf(2), IdOf(3), IdOf(4), IdOf(5));
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

    /// <summary>
    /// Newest-first page bounded by a receivedDateTime cursor. The caller keeps
    /// the ids seen at the boundary timestamp so ties are never lost or repeated.
    /// </summary>
    public async Task<ListPage> ListMessagesAsync(string scope, bool unreadOnly, DateTimeOffset cursor, bool exclusive, int top, CancellationToken ct)
    {
        var path = scope == "inbox" ? "/me/mailFolders/inbox/messages" : "/me/messages";
        var filter = BuildFilter(unreadOnly, cursor, exclusive);
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
                items.Add(new MessageRef(id, when, MessageMetadata.ReadCategories(message),
                    MessageMetadata.GetString(message, "parentFolderId"),
                    message.TryGetProperty("isRead", out var isRead) && isRead.ValueKind == JsonValueKind.True));
            }
        }
        return new ListPage(items, items.Count >= top);
    }

    public static string BuildFilter(bool unreadOnly, DateTimeOffset cursor, bool exclusive = false)
    {
        var stamp = cursor.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);
        var filter = $"receivedDateTime {(exclusive ? "lt" : "le")} {stamp}";
        if (unreadOnly) filter += " and isRead eq false";
        return filter;
    }

    // ----- Batched reads --------------------------------------------------------

    /// <summary>Read many messages through Graph JSON batching; result order matches <paramref name="ids"/>.</summary>
    public async Task<List<BatchPart>> ReadMessagesAsync(IReadOnlyList<string> ids, ReadMode mode, CancellationToken ct)
    {
        var select = mode switch { ReadMode.Metadata => MetadataSelect, ReadMode.Full => FullSelect, _ => "id,categories" };
        var headers = mode == ReadMode.Full
            ? new Dictionary<string, string> { ["Prefer"] = "outlook.body-content-type=\"text\"" }
            : null;
        var requests = ids.Select((id, index) =>
            new BatchRequest(index.ToString(CultureInfo.InvariantCulture), "GET", $"/me/messages/{Uri.EscapeDataString(id)}?$select={select}", headers)).ToList();
        return await SendBatchesAsync(requests, ct);
    }

    // ----- Categories -------------------------------------------------------------

    /// <summary>
    /// Make sure every category exists in the master list (so it has a colour in
    /// Outlook). Returns the stored display names keyed by lower-case name.
    /// </summary>
    public async Task<Dictionary<string, string>> EnsureMasterCategoriesAsync(IReadOnlyList<(string Name, string Color)> wanted, CancellationToken ct)
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

        foreach (var (name, color) in wanted)
        {
            if (existing.ContainsKey(name)) continue;
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
        foreach (var (name, _) in wanted) result[name.ToLowerInvariant()] = existing[name];
        return result;
    }

    // ----- Writes ---------------------------------------------------------------------

    /// <summary>Replace the category list of each message (idempotent; a vanished message counts as done).</summary>
    public Task<WriteOutcome> PatchCategoriesAsync(IReadOnlyList<(string Id, IReadOnlyList<string> Categories)> updates, CancellationToken ct)
    {
        var requests = updates.Select((u, index) => new BatchRequest(
            index.ToString(CultureInfo.InvariantCulture), "PATCH", $"/me/messages/{Uri.EscapeDataString(u.Id)}",
            new Dictionary<string, string> { ["Content-Type"] = "application/json" },
            new JsonObject { ["categories"] = new JsonArray(u.Categories.Select(c => (JsonNode)JsonValue.Create(c)!).ToArray()) })).ToList();
        return ExecuteWritesAsync(requests, "update categories", ct);
    }

    /// <summary>Move messages to the well-known Archive folder (the Outlook equivalent of removing INBOX).</summary>
    public Task<WriteOutcome> MoveToArchiveAsync(IReadOnlyList<string> ids, CancellationToken ct)
    {
        var requests = ids.Select((id, index) => new BatchRequest(
            index.ToString(CultureInfo.InvariantCulture), "POST", $"/me/messages/{Uri.EscapeDataString(id)}/move",
            new Dictionary<string, string> { ["Content-Type"] = "application/json" },
            new JsonObject { ["destinationId"] = "archive" })).ToList();
        return ExecuteWritesAsync(requests, "archive", ct);
    }

    private async Task<WriteOutcome> ExecuteWritesAsync(List<BatchRequest> requests, string action, CancellationToken ct)
    {
        if (requests.Count == 0) return WriteOutcome.Success;
        var pending = requests;
        for (var attempt = 0; ; attempt++)
        {
            var parts = await SendBatchesAsync(pending, ct);
            var retry = new List<BatchRequest>();
            var retryAfter = AppConstants.GraphRetryDelayMs;
            for (var i = 0; i < pending.Count; i++)
            {
                var part = parts[i];
                if (part.Ok || part.Status == 404) continue; // 404: message moved/deleted meanwhile — nothing left to do
                if (part.AuthFailure)
                    return new WriteOutcome(false, false, "Microsoft Graph authorization failed while trying to " + action + " messages. Sign in again (jevoutlook auth).");
                if (!part.Retryable)
                    return new WriteOutcome(false, false, $"Microsoft Graph rejected a request to {action} a message with HTTP {part.Status}. {part.Error}".Trim());
                retryAfter = Math.Max(retryAfter, part.RetryAfterMs);
                retry.Add(pending[i]);
            }
            if (retry.Count == 0) return WriteOutcome.Success;
            if (attempt >= AppConstants.NetworkRetries)
            {
                return new WriteOutcome(false, true,
                    $"Microsoft Graph is temporarily rate-limiting or unavailable ({retry.Count} {action} request(s) failed).");
            }
            await Task.Delay(Math.Min(AppConstants.MaxGraphRetryDelayMs, retryAfter), ct);
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
            if (throwOnError && attempt < AppConstants.NetworkRetries && IsTransient(status))
            {
                var delay = Math.Min(AppConstants.MaxGraphRetryDelayMs, Math.Max(AppConstants.GraphRetryDelayMs, RetryAfterMs(response.Headers)));
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

public sealed class GraphRequestException(HttpStatusCode status, string message) : Exception(message)
{
    public HttpStatusCode Status { get; } = status;

    /// <summary>Rate limiting, server errors and transport failures: safe to pause and continue later.</summary>
    public bool IsTransient => GraphMailClient.IsTransient((int)Status);
}
