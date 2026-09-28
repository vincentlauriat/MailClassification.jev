namespace JevOutlook.Jev;

/// <summary>
/// Outcome of one Decisions request. Failures carry a scope: <c>message</c>
/// (contract violation for this one email) or <c>session</c> (provider-wide
/// problem such as a rate limit, an auth error or a transport failure).
/// </summary>
public sealed class JevResult
{
    public bool Ok { get; init; }
    public string Code { get; init; } = string.Empty;
    public string Error { get; init; } = string.Empty;
    public bool Retryable { get; init; } = true;
    public string FailureScope { get; init; } = "message";
    public int HttpStatus { get; init; }
    public int RetryAfterMs { get; init; }

    public string? RuleId { get; init; }
    public double Confidence { get; init; }
    public IReadOnlyDictionary<string, double>? Probabilities { get; init; }
    public string ProbabilityArgmax { get; init; } = string.Empty;
    public bool ChoiceDiffersFromArgmax { get; init; }

    /// <summary>Cost charged against the run budget for this attempt.</summary>
    public double CostUsd { get; init; }
    /// <summary><c>reported</c>, <c>input_tokens</c> or <c>estimated</c>.</summary>
    public string CostSource { get; init; } = "estimated";
    public long InputTokens { get; init; }
    public string Model { get; init; } = string.Empty;
    public string Provider { get; init; } = string.Empty;
}
