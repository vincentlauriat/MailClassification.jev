namespace JevOutlook;

/// <summary>
/// Tunables shared by the whole application. Values mirror the jevMail (Gmail)
/// reference implementation unless a Microsoft Graph limit dictates otherwise.
/// </summary>
public static class AppConstants
{
    public const string AppName = "jevOutlook";

    // ----- Model providers ------------------------------------------------

    /// <summary>OpenRouter Decisions endpoint (Jev is exposed through it).</summary>
    public const string OpenRouterDecisionsUrl = "https://openrouter.ai/api/alpha/decisions";
    public const string OpenRouterModel = "~typesafe/jev-latest";
    public const string OpenRouterKeysUrl = "https://openrouter.ai/workspaces/default/keys";

    /// <summary>Direct TypeSafe evaluation endpoint (alternative provider).</summary>
    public const string TypeSafeUrl = "https://api.typesafe.ai/v1/systemone";
    public const string TypeSafeModel = "jev-latest";

    /// <summary>Used only for a conservative PRE-request budget estimate.</summary>
    public const double InputRateUsdPerMillion = 0.042;

    /// <summary>Deliberately conservative token estimator.</summary>
    public const double EstimatedTokensPerChar = 1.25;

    /// <summary>Maximum body text sent to Jev on full-body fallback.</summary>
    public const int MaxBodyChars = 5000;

    /// <summary>Hard cap for one Decisions request payload.</summary>
    public const int MaxPayloadBytes = 29000;

    // ----- Microsoft Graph ------------------------------------------------

    public const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    /// <summary>Delegated permissions requested at sign-in.</summary>
    public static readonly string[] GraphScopes =
    [
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/MailboxSettings.ReadWrite",
        "https://graph.microsoft.com/User.Read",
    ];

    /// <summary>Graph JSON batching accepts at most 20 requests per call.</summary>
    public const int GraphBatchRequestLimit = 20;

    /// <summary>
    /// Adaptive read window expressed in messages in flight. Exchange Online enforces a
    /// MailboxConcurrency limit of 4 requests per mailbox and Graph already fans a $batch
    /// out 4 at a time, so exactly one $batch (20 requests) is in flight at any moment —
    /// two concurrent batches were answered with 429 "ApplicationThrottled" (measured
    /// 2026-09-24, Retry-After 8–9 s). The window therefore never exceeds one batch.
    /// </summary>
    public const int GraphInitialConcurrency = 20;
    public const int GraphMinConcurrency = 5;
    public const int GraphMaxConcurrency = 20;
    public const int GraphConcurrencyGrowth = 5;
    public const int GraphMaxConcurrentCalls = 1;
    public const int GraphRetryDelayMs = 500;
    /// <summary>Reads honour the server's Retry-After up to this bound (Exchange often asks for 10–30 s).</summary>
    public const int MaxGraphRetryDelayMs = 30000;
    /// <summary>Transport-level retries of a single call.</summary>
    public const int NetworkRetries = 1;
    /// <summary>Throttled reads inside a batch are retried this many times (window halved each time).</summary>
    public const int GraphReadRetries = 3;

    /// <summary>
    /// Writes (category PATCH, archive move) are throttled much harder than reads by
    /// Exchange Online: batches are sent one at a time and retried with the server's
    /// Retry-After (exponential fallback), for longer than reads.
    /// </summary>
    public const int GraphWriteRetries = 5;
    public const int MaxGraphWriteRetryDelayMs = 30000;

    // ----- Jev waves ------------------------------------------------------

    public const int BatchSize = 50;
    public const int JevInitialConcurrency = 25;
    public const int JevMinConcurrency = 5;
    public const int JevMaxConcurrency = 50;
    public const int JevConcurrencyGrowth = 5;
    public const int JevRetryDelayMs = 400;
    public const int MaxJevRetryDelayMs = 2000;
    public const int MaxConsecutiveJevFailures = 3;
    public const double ProbabilitySumTolerance = 0.0100001;
    public const int MaxSkippedMessageIds = 250;

    // ----- Outlook categories --------------------------------------------

    /// <summary>Outlook master-category colour presets (preset0..preset24).</summary>
    public static readonly string[] CategoryColorPresets =
    [
        "preset0", "preset1", "preset2", "preset3", "preset4", "preset5",
        "preset6", "preset7", "preset8", "preset9", "preset10", "preset11",
        "preset12", "preset13", "preset14", "preset15", "preset16", "preset17",
        "preset18", "preset19", "preset20", "preset21", "preset22", "preset23", "preset24",
    ];

    // ----- Rules ------------------------------------------------------------

    public const int MaxRules = 12;
    public const int MaxRuleName = 50;
    public const int MaxRuleDescription = 400;

    // ----- Run option defaults (same as the jevMail UI) ---------------------

    public const int DefaultLimit = 10;
    public const double DefaultMaxSpendUsd = 0.10;
    public const double DefaultMetadataThreshold = 0.75;
    public const double DefaultArchiveThreshold = 0.93;
}
