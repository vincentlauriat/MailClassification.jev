using System.Globalization;
using MailClassification.Triage;

namespace MailClassification.Cli;

/// <summary>Terminal renderer: timestamped events, a results table and a progress line per batch.</summary>
public sealed class ConsoleSink : ITriageSink
{
    private readonly object _gate = new();
    private bool _headerPrinted;

    public void Event(string level, string message)
    {
        lock (_gate)
        {
            var stamp = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            var tag = level switch { "warn" => "WARN ", "error" => "ERROR", _ => "INFO " };
            var writer = level == "error" ? Console.Error : Console.Out;
            writer.WriteLine($"[{stamp}] {tag} {message}");
        }
    }

    public void Result(ResultRow row)
    {
        lock (_gate)
        {
            if (!_headerPrinted)
            {
                Console.WriteLine();
                Console.WriteLine($"{"From",-32} {"Subject",-44} {"Category",-24} {"Conf.",6} {"Stage",-18} Action");
                Console.WriteLine(new string('-', 150));
                _headerPrinted = true;
            }
            Console.WriteLine($"{Fit(row.From, 32)} {Fit(row.Subject, 44)} {Fit(row.Label, 24)} {row.Confidence,6} {Fit(row.Stage, 18)} {row.Action}");
        }
    }

    public void BatchCompleted(TriageJob job)
    {
        lock (_gate)
        {
            _headerPrinted = false;
            Console.WriteLine();
            Console.WriteLine(ProgressLine(job));
            Console.WriteLine();
        }
    }

    public static string ProgressLine(TriageJob job)
    {
        var target = job.Target > 0 ? job.Target.ToString(CultureInfo.InvariantCulture) : "?";
        return $"Progress: {job.HandledCount}/{target} handled · {job.Processed} categorized ({job.MetadataOnly} metadata, {job.FullBody} full) · " +
               $"{job.Archived} archived · {job.Skipped} skipped · spend ${job.SpentUsd.ToString("0.0000", CultureInfo.InvariantCulture)} / ${job.MaxSpendUsd.ToString("0.00", CultureInfo.InvariantCulture)} · " +
               $"{job.ModelRequests} model requests · {FormatElapsed(job.ElapsedNowMs())}";
    }

    public static void PrintSummary(TriageJob job)
    {
        Console.WriteLine();
        Console.WriteLine($"Session {job.Id}");
        Console.WriteLine($"  Status        : {job.Status}{(job.StopReason.Length > 0 ? " (" + job.StopReason + ")" : string.Empty)}");
        if (job.LastError.Length > 0) Console.WriteLine($"  Last error    : {job.LastError}");
        Console.WriteLine($"  Account       : {job.SignedInAs}");
        Console.WriteLine($"  Scope         : {job.Scope}{(job.UnreadOnly ? ", unread only" : string.Empty)} · limit {(job.Limit is { } l ? l.ToString(CultureInfo.InvariantCulture) : "all")}");
        Console.WriteLine($"  Mode          : {(job.DryRun ? "PREVIEW (no Outlook changes)" : "LIVE")} · {(job.Mode == RunMode.LabelsArchive ? "categories + archive" : "categories only")}");
        Console.WriteLine($"  Thresholds    : metadata ≥ {job.MetadataThreshold.ToString("0.00", CultureInfo.InvariantCulture)} · archive ≥ {job.ArchiveThreshold.ToString("0.00", CultureInfo.InvariantCulture)}");
        Console.WriteLine($"  Messages      : {job.Processed} processed ({job.MetadataOnly} metadata-only, {job.FullBody} full-content), {job.Archived} archived, {job.Skipped} skipped, {job.Failed} failed");
        Console.WriteLine($"  Model         : {job.ModelRequests} requests · {job.ProviderRetries} provider retries · {job.ModelResponseSkips} invalid-response skips");
        Console.WriteLine($"  Cost          : ${job.SpentUsd.ToString("0.0000", CultureInfo.InvariantCulture)} of ${job.MaxSpendUsd.ToString("0.00", CultureInfo.InvariantCulture)} " +
                          $"(reported ${job.ReportedCostUsd.ToString("0.0000", CultureInfo.InvariantCulture)}, from tokens ${job.TokenCalculatedCostUsd.ToString("0.0000", CultureInfo.InvariantCulture)}, estimated ${job.EstimatedCostUsd.ToString("0.0000", CultureInfo.InvariantCulture)})");
        Console.WriteLine($"  Active time   : {FormatElapsed(job.ElapsedNowMs())}");
        if (job.LabelCounts.Count > 0)
        {
            Console.WriteLine("  Categories    :");
            foreach (var label in job.RuleLabels)
            {
                var count = job.LabelCounts.GetValueOrDefault(label.Id);
                Console.WriteLine($"    {Fit(label.Name, 28)} {count,6}{(label.Spam ? "   (archive eligible)" : string.Empty)}");
            }
        }
    }

    public static string FormatElapsed(long ms)
    {
        var t = TimeSpan.FromMilliseconds(ms);
        return t.TotalHours >= 1 ? $"{(int)t.TotalHours}h {t.Minutes:00}m {t.Seconds:00}s" : $"{t.Minutes}m {t.Seconds:00}s";
    }

    private static string Fit(string? value, int width)
    {
        var text = (value ?? string.Empty).Replace('\n', ' ').Replace('\r', ' ');
        if (text.Length > width) text = text[..(width - 1)] + "…";
        return text.PadRight(width);
    }
}
