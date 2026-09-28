using System.Globalization;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using MailClassification.Rules;

namespace MailClassification.Jev;

/// <summary>
/// Sends independent Decisions requests (one per email) and validates the
/// Choice contract strictly: unknown options, missing probabilities or a
/// distribution that does not sum to 1 are rejected without touching Outlook.
/// </summary>
public sealed class JevClient
{
    private readonly HttpClient _http;
    private readonly string _endpoint;
    private readonly string _apiKey;

    public JevClient(HttpClient http, string endpoint, string apiKey)
    {
        _http = http;
        _endpoint = endpoint;
        _apiKey = apiKey;
    }

    public async Task<JevResult> DecideAsync(string payloadText, double estimatedCost, IReadOnlyList<LabelRule> rules, CancellationToken ct)
    {
        HttpResponseMessage response;
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _endpoint);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            request.Content = new StringContent(payloadText, Encoding.UTF8, "application/json");
            response = await _http.SendAsync(request, ct);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested)
        {
            return new JevResult
            {
                Ok = false,
                Code = "transport_error",
                Retryable = true,
                FailureScope = "session",
                CostUsd = Math.Max(0, estimatedCost),
                CostSource = "estimated",
                Error = "The model provider could not be reached. The request will be retried once before processing is paused. Technical detail: " + ex.Message,
            };
        }

        using (response)
        {
            var text = await response.Content.ReadAsStringAsync(ct);
            var retryAfter = response.Headers.RetryAfter is { } header
                ? header.Delta is { } delta ? (int)Math.Max(0, delta.TotalMilliseconds)
                    : header.Date is { } date ? (int)Math.Max(0, (date - DateTimeOffset.UtcNow).TotalMilliseconds) : 0
                : 0;
            return Parse((int)response.StatusCode, retryAfter, text, rules, estimatedCost);
        }
    }

    public static JevResult Parse(int status, int retryAfterMs, string text, IReadOnlyList<LabelRule> rules, double fallbackEstimatedCost)
    {
        var estimated = Math.Max(0, fallbackEstimatedCost);
        JsonElement json = default;
        var hasJson = false;
        try
        {
            using var doc = JsonDocument.Parse(text);
            json = doc.RootElement.Clone();
            hasJson = json.ValueKind == JsonValueKind.Object;
        }
        catch (JsonException) { }

        // ----- cost accounting: reported cost > input tokens > estimate -----
        double actualCost = double.NaN;
        long inputTokens = -1;
        string model = string.Empty, provider = string.Empty;
        if (hasJson)
        {
            model = ReadString(json, "model");
            provider = ReadString(json, "provider");
            if (json.TryGetProperty("usage", out var usage) && usage.ValueKind == JsonValueKind.Object)
            {
                actualCost = ReadNumber(usage, "cost") ?? ReadNumber(usage, "total_cost") ?? ReadNumber(usage, "totalCost") ?? double.NaN;
                var tokens = ReadNumber(usage, "input_tokens") ?? ReadNumber(usage, "prompt_tokens");
                if (tokens is { } t && t >= 0) inputTokens = (long)t;
            }
        }
        var hasReportedCost = !double.IsNaN(actualCost) && actualCost >= 0;
        var hasInputTokens = inputTokens >= 0;
        // A 429 is rejected before inference, so nothing was billed; a timeout or 5xx is
        // uncertain and keeps the conservative reservation.
        var cost = hasReportedCost ? actualCost
            : hasInputTokens ? inputTokens * AppConstants.InputRateUsdPerMillion / 1_000_000d
            : status == 429 ? 0
            : estimated;
        var costSource = hasReportedCost ? "reported" : hasInputTokens ? "input_tokens" : "estimated";

        JevResult Failure(string code, string message, bool retryable = true, string scope = "message") => new()
        {
            Ok = false,
            Code = code,
            Error = message,
            Retryable = retryable,
            FailureScope = scope,
            HttpStatus = status,
            RetryAfterMs = Math.Min(AppConstants.MaxJevRetryDelayMs, Math.Max(0, retryAfterMs)),
            CostUsd = cost,
            CostSource = costSource,
            InputTokens = hasInputTokens ? inputTokens : 0,
            Model = model,
            Provider = provider,
        };

        if (status < 200 || status >= 300)
        {
            // Never echo arbitrary provider text that might contain request data.
            if (status == 401) return Failure("invalid_api_key", "The model provider rejected the API key. Check the key and try again.", false, "session");
            if (status == 402) return Failure("insufficient_credits", "The provider account has insufficient credits. Add credits and try again.", false, "session");
            if (status == 403) return Failure("access_denied", "The provider denied access to the requested model. Check the key permissions and workspace limits.", false, "session");
            if (status is 408 or 409 or 425 or 429 || status >= 500)
                return Failure("provider_temporarily_unavailable", $"The model provider temporarily returned HTTP {status}. No Outlook changes were made.", true, "session");
            return Failure("request_rejected", $"The model provider rejected the Decisions request with HTTP {status}. No Outlook changes were made.", false, "session");
        }

        if (!hasJson) return Failure("invalid_json", "The provider returned a response that was not valid JSON. No Outlook changes were made.");

        if (!json.TryGetProperty("answers", out var answers) || answers.ValueKind != JsonValueKind.Object ||
            !answers.TryGetProperty(JevPayloadBuilder.QuestionId, out var answer) || answer.ValueKind != JsonValueKind.Object)
        {
            return Failure("missing_answer", "Jev returned no classification answer. No Outlook changes were made.");
        }

        var choice = ReadString(answer, "choice");
        var rule = OptionIndex(choice) is { } idx && idx < rules.Count ? rules[idx] : null;
        if (rule is null) return Failure("invalid_choice", "Jev returned a category outside the configured classification rules. No Outlook changes were made.");

        // Jev confidence measures distribution separation and is NOT the selected
        // option probability. Use the documented confidence for both thresholds.
        var confidence = ReadNumber(answer, "confidence");
        if (confidence is null || double.IsNaN(confidence.Value) || double.IsInfinity(confidence.Value) || confidence < 0 || confidence > 1)
            return Failure("invalid_confidence", "Jev returned an invalid confidence value. No Outlook changes were made.");

        if (!answer.TryGetProperty("probabilities", out var probabilities) || probabilities.ValueKind != JsonValueKind.Object)
            return Failure("invalid_probabilities", "Jev returned an invalid probability distribution. No Outlook changes were made.");

        var expected = new HashSet<string>(Enumerable.Range(0, rules.Count).Select(i => "L" + i), StringComparer.Ordinal);
        var returned = probabilities.EnumerateObject().Select(p => p.Name).ToList();
        var missing = expected.Count(k => !returned.Contains(k));
        var unexpected = returned.Count(k => !expected.Contains(k));
        if (missing > 0)
            return Failure("missing_probabilities", $"Jev omitted {missing} of {expected.Count} configured class probabilities. No Outlook changes were made.");
        if (unexpected > 0)
            return Failure("unexpected_probability_keys", $"Jev returned {unexpected} unexpected probability key{(unexpected == 1 ? string.Empty : "s")}. No Outlook changes were made.");

        var distribution = new Dictionary<string, double>(StringComparer.Ordinal);
        double total = 0, maxProbability = -1;
        var maxKey = string.Empty;
        for (var i = 0; i < rules.Count; i++)
        {
            var key = "L" + i;
            var p = ReadNumber(probabilities, key);
            if (p is null || double.IsNaN(p.Value) || double.IsInfinity(p.Value) || p < 0 || p > 1)
                return Failure("invalid_probability_value", "Jev returned an invalid class probability. No Outlook changes were made.");
            distribution[key] = p.Value;
            if (p > maxProbability) { maxProbability = p.Value; maxKey = key; }
            total += p.Value;
        }
        if (Math.Abs(total - 1) > AppConstants.ProbabilitySumTolerance)
            return Failure("probability_sum_mismatch", $"Jev returned class probabilities with a total of {Math.Round(total, 6).ToString(CultureInfo.InvariantCulture)} instead of 1. No Outlook changes were made.");

        return new JevResult
        {
            Ok = true,
            HttpStatus = status,
            RuleId = rule.Id,
            Confidence = confidence.Value,
            Probabilities = distribution,
            ProbabilityArgmax = maxKey,
            ChoiceDiffersFromArgmax = maxKey != choice,
            CostUsd = cost,
            CostSource = costSource,
            InputTokens = hasInputTokens ? inputTokens : 0,
            Model = model,
            Provider = provider,
        };
    }

    private static int? OptionIndex(string choice)
    {
        if (choice.Length < 2 || choice[0] != 'L') return null;
        var digits = choice[1..];
        if (digits.Length > 1 && digits[0] == '0') return null;
        return int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var index) ? index : null;
    }

    private static string ReadString(JsonElement element, string property) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static double? ReadNumber(JsonElement element, string property)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(property, out var value)) return null;
        return value.ValueKind switch
        {
            JsonValueKind.Number when value.TryGetDouble(out var d) => d,
            JsonValueKind.String when double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => null,
        };
    }
}
