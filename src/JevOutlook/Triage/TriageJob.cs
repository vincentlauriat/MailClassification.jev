using System.Text.Json.Serialization;

namespace JevOutlook.Triage;

public static class JobStatus
{
    public const string Running = "running";
    public const string Paused = "paused";
    public const string Completed = "completed";
    public const string Budget = "budget";
    public const string Cancelled = "cancelled";
    public const string Error = "error";
}

public static class RunMode
{
    public const string Labels = "labels";
    public const string LabelsArchive = "labels_archive";
}

public sealed class Decision
{
    public string RuleId { get; set; } = string.Empty;
    public double Confidence { get; set; }
    /// <summary><c>metadata</c>, <c>metadata_fallback</c> or <c>full</c>.</summary>
    public string Stage { get; set; } = "metadata";
}

public sealed class PendingItem
{
    public string Id { get; set; } = string.Empty;
    public Decision? MetadataResult { get; set; }
    public Decision? Final { get; set; }
}

public sealed record RuleLabelSummary(string Id, string Name, bool Spam);

/// <summary>
/// Persisted state of one processing session. Every paid decision is saved
/// before Outlook is modified, so a run can be paused, resumed or replayed
/// without sending the same message to the model twice.
/// </summary>
public sealed class TriageJob
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public long CreatedAt { get; set; }
    public long UpdatedAt { get; set; }
    public long ActiveStartedAt { get; set; }
    public long ElapsedMs { get; set; }
    public long FinishedAt { get; set; }
    public string Status { get; set; } = JobStatus.Running;
    public string SignedInAs { get; set; } = string.Empty;
    public string AccountId { get; set; } = string.Empty;
    /// <summary>"Outlook", "Gmail" or "IMAP" — for wording only.</summary>
    public string Provider { get; set; } = string.Empty;

    // Options frozen at start
    public string Scope { get; set; } = "inbox";
    public bool UnreadOnly { get; set; } = true;
    public bool DryRun { get; set; } = true;
    public string Mode { get; set; } = RunMode.Labels;
    public double MaxSpendUsd { get; set; }
    public double MetadataThreshold { get; set; }
    public double ArchiveThreshold { get; set; }
    /// <summary>null = all matching messages.</summary>
    public int? Limit { get; set; }

    // Listing cursor (newest-first): the provider's sort key (Graph: ISO timestamp; IMAP: zero-padded UID)
    public string Cursor { get; set; } = string.Empty;
    public List<string> CursorBoundaryIds { get; set; } = [];
    /// <summary>True when a whole page shared the cursor timestamp: the next query uses <c>lt</c> instead of <c>le</c>.</summary>
    public bool CursorExclusive { get; set; }
    public bool Exhausted { get; set; }

    // Counters
    public long InitialEstimate { get; set; }
    public long Target { get; set; }
    public int Processed { get; set; }
    public int MetadataOnly { get; set; }
    public int FullBody { get; set; }
    public int Archived { get; set; }
    public int Failed { get; set; }
    public int Skipped { get; set; }
    public int ProviderRetries { get; set; }
    public int ModelResponseSkips { get; set; }
    public int ConsecutiveJevFailures { get; set; }
    public int ModelRequests { get; set; }
    public int GraphRateLimitRetries { get; set; }
    public List<string> SkippedMessageIds { get; set; } = [];

    // Cost
    public double SpentUsd { get; set; }
    public double ReportedCostUsd { get; set; }
    public double TokenCalculatedCostUsd { get; set; }
    public double EstimatedCostUsd { get; set; }

    // Adaptive windows
    public int GraphConcurrency { get; set; } = AppConstants.GraphInitialConcurrency;
    public int JevConcurrency { get; set; } = AppConstants.JevInitialConcurrency;

    public List<RuleLabelSummary> RuleLabels { get; set; } = [];
    public Dictionary<string, int> LabelCounts { get; set; } = new(StringComparer.Ordinal);
    public List<PendingItem> Pending { get; set; } = [];
    public string StopReason { get; set; } = string.Empty;
    public string LastError { get; set; } = string.Empty;

    [JsonIgnore] public int HandledCount => Processed + Skipped;
    [JsonIgnore] public bool IsRunning => Status == JobStatus.Running;

    public long ElapsedNowMs()
    {
        var active = ActiveStartedAt > 0 ? Math.Max(0, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() - ActiveStartedAt) : 0;
        return ElapsedMs + active;
    }
}
