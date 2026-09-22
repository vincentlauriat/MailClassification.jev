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
    /// Adaptive read window expressed in messages in flight. Every request inside a
    /// $batch counts individually against Outlook throttling (about 4 concurrent
    /// requests per app × mailbox, and Graph itself fans a batch out 4 at a time), so
    /// at most <see cref="GraphMaxConcurrentCalls"/> batch calls run in parallel and the
    /// window never exceeds 2 batches of 20.
    /// </summary>
    public const int GraphInitialConcurrency = 20;
    public const int GraphMinConcurrency = 5;
    public const int GraphMaxConcurrency = 40;
    public const int GraphConcurrencyGrowth = 20;
    public const int GraphMaxConcurrentCalls = 2;
    public const int GraphRetryDelayMs = 500;
    public const int MaxGraphRetryDelayMs = 3000;
    public const int NetworkRetries = 1;

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

    /// <summary>Technical category applied to successfully processed messages.</summary>
    public const string TechnicalTriagedCategory = "jev-triaged";

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
