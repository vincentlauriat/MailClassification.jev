using System.Globalization;
using System.Text.Json;
using JevOutlook.Graph;
using JevOutlook.Jev;
using JevOutlook.Rules;
using JevOutlook.Storage;

namespace JevOutlook.Triage;

/// <summary>
/// Two-stage, confidence-aware classification of Outlook messages:
/// 1) classify every message from metadata (headers + preview);
/// 2) classify only low-confidence / archive-sensitive messages from full content.
/// Every model decision is checkpointed before Outlook is modified, and Outlook
/// writes are idempotent so a replay after an interruption never pays twice.
/// </summary>
public sealed class TriageEngine
{
    private readonly GraphMailClient _graph;
    private readonly JevClient _jev;
    private readonly string _model;
    private readonly ITriageSink _sink;

    public TriageEngine(GraphMailClient graph, JevClient jev, string model, ITriageSink sink)
    {
        _graph = graph;
        _jev = jev;
        _model = model;
        _sink = sink;
    }

    // =====================================================================
    // Session lifecycle
    // =====================================================================

    public async Task<TriageJob> StartAsync(RunOptions options, IReadOnlyList<LabelRule> rules, CancellationToken ct)
    {
        var existing = JobStore.Load();
        if (existing is { Status: JobStatus.Running })
        {
            throw new InvalidOperationException("A processing session is already active. Stop or continue it before starting another (jevoutlook status).");
        }

        var normalizedRules = RuleStore.Save(rules);
        var signedInAs = await _graph.GetSignedInUserAsync(ct);
        var folders = await _graph.GetWellKnownFoldersAsync(ct);
        if (!options.DryRun && options.Mode == RunMode.LabelsArchive && folders.Archive is null)
        {
            throw new InvalidOperationException("This mailbox has no Archive folder, so archive mode cannot be used. Run in labels-only mode.");
        }
        if (!options.DryRun)
        {
            await EnsureCategoriesAsync(normalizedRules, ct);
        }

        var estimate = await _graph.EstimateAsync(options.Scope, options.UnreadOnly, ct);
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var job = new TriageJob
        {
            CreatedAt = now,
            UpdatedAt = now,
            SignedInAs = signedInAs,
            Scope = options.Scope,
            UnreadOnly = options.UnreadOnly,
            DryRun = options.DryRun,
            Mode = options.Mode,
            MaxSpendUsd = options.MaxSpendUsd,
            MetadataThreshold = options.MetadataThreshold,
            ArchiveThreshold = options.ArchiveThreshold,
            Limit = options.Limit,
            Cursor = DateTimeOffset.UtcNow.AddMinutes(5),
            InitialEstimate = estimate,
            Target = options.Limit is { } limit
                ? (estimate >= 0 ? Math.Min(limit, estimate) : limit)
                : Math.Max(0, estimate),
            RuleLabels = normalizedRules.Select(r => new RuleLabelSummary(r.Id, r.Name, r.Spam)).ToList(),
        };
        foreach (var rule in normalizedRules) job.LabelCounts[rule.Id] = 0;

        RuleStore.SaveJobRules(normalizedRules);
        JobStore.Save(job);

        var found = estimate >= 0 ? $"Found about {estimate} matching messages." : "The number of matching messages is unknown.";
        _sink.Event("info", $"{found} Messages selected for processing: {(options.Limit is { } l ? l.ToString(CultureInfo.InvariantCulture) : "all")}.");
        return job;
    }

    public static TriageJob Resume(double? newMaxSpend)
    {
        var job = JobStore.Load() ?? throw new InvalidOperationException("There is no processing session to continue. Start one with: jevoutlook run");
        if (job.Status is JobStatus.Completed or JobStatus.Cancelled)
            throw new InvalidOperationException("This processing session is already finished. Start a new session to continue.");
        if (job.Status == JobStatus.Budget)
        {
            if (newMaxSpend is null || newMaxSpend <= job.MaxSpendUsd)
                throw new InvalidOperationException(
                    $"This processing session reached its cost limit (${job.MaxSpendUsd:0.00}). Continue with a higher limit: jevoutlook continue --max-spend <usd>");
        }
        if (newMaxSpend is { } spend)
        {
            if (spend <= 0 || spend > 1000) throw new ArgumentException("The maximum processing cost must be greater than $0 and no more than $1000.");
            job.MaxSpendUsd = spend;
        }
        job.Status = JobStatus.Running;
        job.StopReason = string.Empty;
        job.LastError = string.Empty;
        job.ConsecutiveJevFailures = 0; // re-arm the circuit breaker: the user chose to try again
        JobStore.Save(job);
        return job;
    }

    /// <summary>Process batches until the session completes, pauses, hits its budget, fails or is cancelled.</summary>
    public async Task RunLoopAsync(TriageJob job, CancellationToken ct)
    {
        var rules = RuleStore.LoadJobRules();
        var folders = await _graph.GetWellKnownFoldersAsync(ct);
        var categoryNames = job.DryRun ? null : await EnsureCategoriesAsync(rules, ct);

        while (job.IsRunning)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                await ProcessNextBatchAsync(job, rules, folders, categoryNames, ct);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                if (job.IsRunning)
                {
                    job.Status = JobStatus.Paused;
                    job.StopReason = "user-stop";
                }
                CloseTiming(job);
                JobStore.Save(job);
                _sink.Event("warn", "Processing stopped. All decisions are checkpointed; run 'jevoutlook continue' to resume.");
                return;
            }
            _sink.BatchCompleted(job);
        }
    }

    private async Task<Dictionary<string, string>> EnsureCategoriesAsync(IReadOnlyList<LabelRule> rules, CancellationToken ct)
    {
        var wanted = new List<(string, string)>();
        for (var i = 0; i < rules.Count; i++)
        {
            wanted.Add((rules[i].Name, AppConstants.CategoryColorPresets[i % AppConstants.CategoryColorPresets.Length]));
        }
        wanted.Add((AppConstants.TechnicalTriagedCategory, "none"));
        return await _graph.EnsureMasterCategoriesAsync(wanted, ct);
    }

    // =====================================================================
    // One batch
    // =====================================================================

    private sealed class WaveRequest
    {
        public required PendingItem Item { get; init; }
        public required MessageMetadata Metadata { get; init; }
        public required string Stage { get; init; }
        public required string PayloadText { get; init; }
        public required double EstimatedCost { get; init; }
        public JevResult? Result { get; set; }
    }

    private sealed record FinalResult(PendingItem Item, LabelRule Rule, double Confidence, string Stage, bool Archive);

    private async Task ProcessNextBatchAsync(TriageJob job, IReadOnlyList<LabelRule> rules, WellKnownFolders folders,
        Dictionary<string, string>? categoryNames, CancellationToken ct)
    {
        job.ActiveStartedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var results = new List<ResultRow>();

        try
        {
            if (job.Pending.Count == 0)
            {
                if (job.Limit is { } limit && job.HandledCount >= limit)
                {
                    Complete(job, "target-reached");
                }
                else
                {
                    var wanted = Math.Min(AppConstants.BatchSize, job.Limit is { } l ? l - job.HandledCount : AppConstants.BatchSize);
                    await FillPendingAsync(job, wanted, folders, ct);
                    if (job.Pending.Count == 0 && job.IsRunning) Complete(job, "no-more-messages");
                }
                JobStore.Save(job);
            }

            if (!job.IsRunning) return;

            // ----- metadata read -------------------------------------------------
            var metadataById = new Dictionary<string, MessageMetadata>(StringComparer.Ordinal);
            var snapshot = job.Pending.ToList();
            var metadataRead = await ReadAdaptiveAsync(snapshot.Select(i => i.Id).ToList(), ReadMode.Metadata, job, ct);

            foreach (var item in snapshot)
            {
                if (!metadataRead.Messages.TryGetValue(item.Id, out var message)) continue;
                var metadata = MessageMetadata.FromGraph(message);
                metadataById[item.Id] = metadata;
                // A pending item with a final decision may already carry the marker if an
                // earlier grouped write partly succeeded before its checkpoint. Keep it and
                // replay the idempotent write so counters can advance.
                if (!job.DryRun && item.Final is null && HasMarker(metadata.Categories))
                {
                    RemovePending(job, item.Id);
                }
            }
            foreach (var id in metadataRead.UnavailableIds)
            {
                var replayed = job.Pending.FirstOrDefault(p => p.Id == id);
                if (!job.DryRun && replayed?.Final is not null)
                {
                    // The message id changed because an earlier archive move succeeded before its
                    // checkpoint. The decision is final and already applied: count it, do not skip it.
                    ConfirmReplayedFinal(job, replayed, rules, results);
                    continue;
                }
                SkipUnavailable(job, id, metadataById.GetValueOrDefault(id), results);
            }

            var budgetBlocked = false;

            // ----- stage 1: metadata wave ------------------------------------------
            if (job.IsRunning)
            {
                var metadataRequests = job.Pending
                    .Where(item => item.Final is null && item.MetadataResult is null)
                    .Select(item => CreateWaveRequest(item, metadataById.GetValueOrDefault(item.Id), rules, "metadata", string.Empty))
                    .ToList();

                budgetBlocked |= await DispatchWaveAsync(metadataRequests, rules, job, ct);

                // Persist every successful answer before processing failures.
                foreach (var request in metadataRequests)
                {
                    if (request.Result is { Ok: true } ok)
                    {
                        request.Item.MetadataResult = new Decision { RuleId = ok.RuleId!, Confidence = ok.Confidence, Stage = "metadata" };
                    }
                }
                JobStore.Save(job);

                foreach (var request in metadataRequests)
                {
                    if (!job.IsRunning) break;
                    if (request.Result is null) continue;
                    if (request.Result.Ok) { job.ConsecutiveJevFailures = 0; continue; }
                    HandleJevFailure(job, request.Item, request.Metadata, "metadata", request.Result, results);
                }
            }

            // ----- stage 2: full-content wave --------------------------------------
            if (job.IsRunning)
            {
                var fullItems = new List<PendingItem>();
                foreach (var item in job.Pending.ToList())
                {
                    if (item.Final is not null || item.MetadataResult is null) continue;
                    var selected = RuleValidator.ById(rules, item.MetadataResult.RuleId)
                        ?? throw new InvalidOperationException("Jev selected an unknown classification rule.");
                    var needsFull = item.MetadataResult.Confidence < job.MetadataThreshold ||
                                    (job.Mode == RunMode.LabelsArchive && selected.Spam && item.MetadataResult.Confidence < job.ArchiveThreshold);
                    if (!needsFull)
                    {
                        item.Final = new Decision { RuleId = item.MetadataResult.RuleId, Confidence = item.MetadataResult.Confidence, Stage = "metadata" };
                        item.MetadataResult = null;
                        continue;
                    }
                    fullItems.Add(item);
                }

                var fullRead = await ReadAdaptiveAsync(fullItems.Select(i => i.Id).ToList(), ReadMode.Full, job, ct);
                foreach (var id in fullRead.UnavailableIds)
                {
                    SkipUnavailable(job, id, metadataById.GetValueOrDefault(id), results);
                }

                var fullRequests = new List<WaveRequest>();
                if (job.IsRunning)
                {
                    foreach (var item in fullItems)
                    {
                        if (!job.Pending.Any(p => p.Id == item.Id)) continue;
                        if (!fullRead.Messages.TryGetValue(item.Id, out var full)) continue;
                        var metadata = metadataById.GetValueOrDefault(item.Id) ?? MessageMetadata.Empty(item.Id);
                        var (text, reason) = MessageContent.Extract(full);
                        var body = MessageContent.Compact(text);

                        if (body.Length == 0)
                        {
                            var fallback = RuleValidator.ReviewFallback(rules);
                            if (fallback is not null)
                            {
                                item.Final = new Decision { RuleId = fallback.Id, Confidence = item.MetadataResult!.Confidence, Stage = "metadata_fallback" };
                                item.MetadataResult = null;
                                _sink.Event("warn",
                                    $"Could not extract full text from “{Clip(metadata.Subject, "(no subject)")}”: {reason} The safe “{fallback.Name}” fallback " +
                                    (job.DryRun ? "would be applied in a live run." : "will be applied without archiving."));
                            }
                            else
                            {
                                SkipUnreadable(job, item, metadata, reason, results);
                            }
                            continue;
                        }

                        fullRequests.Add(CreateWaveRequest(item, metadata, rules, "full", body));
                    }
                }
                JobStore.Save(job);

                if (job.IsRunning && fullRequests.Count > 0)
                {
                    budgetBlocked |= await DispatchWaveAsync(fullRequests, rules, job, ct);
                    foreach (var request in fullRequests)
                    {
                        if (request.Result is { Ok: true } ok)
                        {
                            request.Item.Final = new Decision { RuleId = ok.RuleId!, Confidence = ok.Confidence, Stage = "full" };
                            request.Item.MetadataResult = null;
                        }
                    }
                    JobStore.Save(job);

                    foreach (var request in fullRequests)
                    {
                        if (!job.IsRunning) break;
                        if (request.Result is null) continue;
                        if (request.Result.Ok) { job.ConsecutiveJevFailures = 0; continue; }
                        HandleJevFailure(job, request.Item, request.Metadata, "full", request.Result, results);
                    }
                }
            }

            // ----- apply the resolved prefix as grouped Outlook writes ----------------
            var readyItems = new List<PendingItem>();
            foreach (var item in job.Pending)
            {
                if (item.Final is null || !metadataById.ContainsKey(item.Id)) break;
                readyItems.Add(item);
            }
            var finalized = readyItems.Select(item =>
            {
                var selected = RuleValidator.ById(rules, item.Final!.RuleId)
                    ?? throw new InvalidOperationException("Jev selected an unknown classification rule.");
                var archive = job.Mode == RunMode.LabelsArchive && selected.Spam && item.Final.Confidence >= job.ArchiveThreshold;
                return new FinalResult(item, selected, item.Final.Confidence, item.Final.Stage, archive);
            }).ToList();

            if (finalized.Count > 0 && !job.DryRun)
            {
                var outcome = await ApplyFinalResultsAsync(finalized, metadataById, categoryNames!, ct);
                if (!outcome.Ok)
                {
                    if (!outcome.Retryable) throw new InvalidOperationException(outcome.Error);
                    job.Status = JobStatus.Paused;
                    job.StopReason = "graph-temporary";
                    job.LastError = outcome.Error;
                    _sink.Event("warn", outcome.Error + " Processing is paused and all decisions are checkpointed. Continue later to retry safely.");
                }
            }

            if (job.DryRun || job.Status != JobStatus.Paused || job.StopReason != "graph-temporary")
            {
                foreach (var final in finalized)
                {
                    job.Processed += 1;
                    job.LabelCounts[final.Rule.Id] = job.LabelCounts.GetValueOrDefault(final.Rule.Id) + 1;
                    if (final.Stage is "metadata" or "metadata_fallback") job.MetadataOnly += 1; else job.FullBody += 1;
                    if (final.Archive && !job.DryRun) job.Archived += 1;

                    var metadata = metadataById.GetValueOrDefault(final.Item.Id) ?? MessageMetadata.Empty(final.Item.Id);
                    var action = final.Stage == "metadata_fallback"
                        ? (job.DryRun ? "would apply review fallback" : "review fallback applied")
                        : final.Archive ? (job.DryRun ? "would archive" : "archived")
                        : (job.DryRun ? "would categorize" : "categorized");
                    results.Add(new ResultRow(metadata.From, metadata.Subject, final.Rule.Name,
                        final.Confidence.ToString("0.000", CultureInfo.InvariantCulture), final.Stage, action));
                }
                if (readyItems.Count > 0) job.Pending.RemoveRange(0, readyItems.Count);
                job.Target = Math.Max(job.Target, job.HandledCount + job.Pending.Count);
                JobStore.Save(job);
            }

            if (budgetBlocked && job.IsRunning)
            {
                job.Status = JobStatus.Budget;
                job.StopReason = "budget";
            }

            if (job.IsRunning && job.Pending.Count == 0)
            {
                if (job.Limit is { } limit && job.HandledCount >= limit) Complete(job, "target-reached");
                else if (job.Exhausted) Complete(job, "no-more-messages");
            }

            if (job.Status == JobStatus.Completed) job.Target = job.HandledCount;
            if (job.Status == JobStatus.Budget)
            {
                _sink.Event("warn", "The cost limit prevents the next model request. Unfinished messages remain unchanged.");
            }
            else if (job.Status is JobStatus.Running or JobStatus.Completed)
            {
                _sink.Event("info", $"Reviewed {results.Count} message(s) in this batch.");
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch (GraphRequestException transient) when (transient.IsTransient)
        {
            job.Status = JobStatus.Paused;
            job.StopReason = "graph-temporary";
            job.LastError = Clip(transient.Message, "Microsoft Graph is temporarily unavailable.", 500);
            _sink.Event("warn", job.LastError + " Processing is paused and all decisions are checkpointed. Continue later to retry safely.");
        }
        catch (Exception error)
        {
            job.Status = JobStatus.Error;
            job.StopReason = "processing-error";
            job.Failed += 1;
            job.LastError = Clip(error.Message, "Unknown error", 500);
            _sink.Event("error", job.LastError);
        }
        finally
        {
            CloseTiming(job);
            JobStore.Save(job);
            foreach (var row in results) _sink.Result(row);
        }
    }

    /// <summary>A checkpointed decision whose message is gone (moved by an earlier partial write) is complete.</summary>
    private void ConfirmReplayedFinal(TriageJob job, PendingItem item, IReadOnlyList<LabelRule> rules, List<ResultRow> results)
    {
        var rule = RuleValidator.ById(rules, item.Final!.RuleId)
            ?? throw new InvalidOperationException("Jev selected an unknown classification rule.");
        var archive = job.Mode == RunMode.LabelsArchive && rule.Spam && item.Final.Confidence >= job.ArchiveThreshold;
        job.Processed += 1;
        job.LabelCounts[rule.Id] = job.LabelCounts.GetValueOrDefault(rule.Id) + 1;
        if (item.Final.Stage is "metadata" or "metadata_fallback") job.MetadataOnly += 1; else job.FullBody += 1;
        if (archive) job.Archived += 1;
        RemovePending(job, item.Id);
        results.Add(new ResultRow(string.Empty, "(message already processed in a previous batch)", rule.Name,
            item.Final.Confidence.ToString("0.000", CultureInfo.InvariantCulture), item.Final.Stage,
            archive ? "archived (confirmed on replay)" : "categorized (confirmed on replay)"));
    }

    private static void Complete(TriageJob job, string reason)
    {
        job.Status = JobStatus.Completed;
        job.StopReason = reason;
    }

    private static void CloseTiming(TriageJob job)
    {
        if (job.ActiveStartedAt > 0)
        {
            job.ElapsedMs += Math.Max(0, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() - job.ActiveStartedAt);
        }
        job.ActiveStartedAt = 0;
        if (!job.IsRunning) job.FinishedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
    }

    // =====================================================================
    // Listing with a receivedDateTime cursor
    // =====================================================================

    /// <summary>
    /// Fill <see cref="TriageJob.Pending"/> with up to <paramref name="wanted"/> eligible
    /// messages. The cursor only advances over messages that were actually examined, so
    /// nothing is lost when a page contains more candidates than needed.
    /// </summary>
    private async Task FillPendingAsync(TriageJob job, int wanted, WellKnownFolders folders, CancellationToken ct)
    {
        var collected = new List<PendingItem>();
        var skippedLookup = new HashSet<string>(job.SkippedMessageIds, StringComparer.Ordinal);
        var guard = 0;

        bool Eligible(MessageRef item) =>
            !(job.Scope == "all" && folders.IsExcluded(item.ParentFolderId)) &&
            !skippedLookup.Contains(item.Id) &&
            !HasMarker(item.Categories);

        while (collected.Count < wanted && !job.Exhausted && guard < MaxScanPagesPerBatch)
        {
            guard += 1;
            var top = Math.Clamp(wanted - collected.Count + job.CursorBoundaryIds.Count + skippedLookup.Count, 20, 200);
            var page = await _graph.ListMessagesAsync(job.Scope, job.UnreadOnly, job.Cursor, job.CursorExclusive, top, ct);
            ExaminePage(job, page, wanted, Eligible, collected);
        }

        job.Pending = collected;

        if (collected.Count == 0 && !job.Exhausted && job.IsRunning)
        {
            // Every page was made of already-processed messages. Pause with a clear reason instead of
            // declaring the mailbox exhausted; 'continue' resumes the scan from the cursor.
            job.Status = JobStatus.Paused;
            job.StopReason = "scan-guard";
            job.LastError = $"Scanned {MaxScanPagesPerBatch} pages of already-processed messages without finding a new one.";
            _sink.Event("warn", job.LastError + " Run 'jevoutlook continue' to keep scanning older mail, or start a new session with a narrower scope.");
        }
    }

    private const int MaxScanPagesPerBatch = 40;

    /// <summary>
    /// Consume one listing page: advance the cursor over examined items only, keep the
    /// ids that share the cursor timestamp, collect eligible messages up to <paramref name="wanted"/>.
    /// Pure function of the job cursor state, exposed for unit tests.
    /// </summary>
    internal static void ExaminePage(TriageJob job, ListPage page, int wanted, Func<MessageRef, bool> eligible, List<PendingItem> collected)
    {
        if (page.Items.Count == 0)
        {
            job.Exhausted = true;
            return;
        }

        var boundary = new HashSet<string>(job.CursorBoundaryIds, StringComparer.Ordinal);
        var examined = 0;
        var stoppedEarly = false;
        foreach (var item in page.Items)
        {
            if (boundary.Contains(item.Id)) continue;
            if (collected.Count >= wanted) { stoppedEarly = true; break; }

            examined += 1;
            if (item.ReceivedDateTime == job.Cursor && !job.CursorExclusive)
            {
                job.CursorBoundaryIds.Add(item.Id);
            }
            else
            {
                job.Cursor = item.ReceivedDateTime;
                job.CursorBoundaryIds = [item.Id];
                job.CursorExclusive = false;
            }

            if (eligible(item) && !collected.Any(c => c.Id == item.Id)) collected.Add(new PendingItem { Id = item.Id });
        }

        if (!stoppedEarly && !page.Full)
        {
            job.Exhausted = true;
        }
        else if (examined == 0 && page.Full)
        {
            // Pathological case: a whole page shares the boundary timestamp. Query strictly
            // older mail next (lt instead of le): nothing in between is lost or repeated; only
            // further messages with that exact same timestamp would be left out.
            job.CursorExclusive = true;
            job.CursorBoundaryIds = [];
        }
    }

    private static bool HasMarker(IEnumerable<string> categories) =>
        categories.Any(c => string.Equals(c, AppConstants.TechnicalTriagedCategory, StringComparison.OrdinalIgnoreCase));

    // =====================================================================
    // Adaptive Graph reads
    // =====================================================================

    private sealed record ReadOutcome(Dictionary<string, JsonElement> Messages, List<string> UnavailableIds);

    private async Task<ReadOutcome> ReadAdaptiveAsync(IReadOnlyList<string> ids, ReadMode mode, TriageJob job, CancellationToken ct)
    {
        var output = new ReadOutcome(new Dictionary<string, JsonElement>(StringComparer.Ordinal), []);
        if (ids.Count == 0 || !job.IsRunning) return output;

        var queue = ids.Select(id => (Id: id, Attempt: 0)).ToList();
        while (queue.Count > 0 && job.IsRunning)
        {
            var window = Math.Min(NormalizeConcurrency(job.GraphConcurrency, AppConstants.GraphMinConcurrency, AppConstants.GraphMaxConcurrency), queue.Count);
            var entries = queue.Take(window).ToList();
            queue.RemoveRange(0, window);

            var parts = await _graph.ReadMessagesAsync(entries.Select(e => e.Id).ToList(), mode, ct);
            var retry = new List<(string Id, int Attempt)>();
            var exhausted = new List<string>();
            var retryAfterMs = AppConstants.GraphRetryDelayMs;
            var adaptiveFailure = false;

            for (var i = 0; i < entries.Count; i++)
            {
                var part = parts[i];
                if (part.Ok) { output.Messages[entries[i].Id] = part.Body; continue; }
                if (part.Status == 404) { output.UnavailableIds.Add(entries[i].Id); continue; }
                if (part.AuthFailure) throw new InvalidOperationException("Microsoft Graph authorization failed while reading messages. Sign in again (jevoutlook auth) and try again.");
                if (!part.Retryable) throw new InvalidOperationException(part.Error.Length > 0 ? part.Error : $"Microsoft Graph rejected a message read with HTTP {part.Status}.");
                adaptiveFailure = true;
                retryAfterMs = Math.Max(retryAfterMs, part.RetryAfterMs);
                if (entries[i].Attempt < AppConstants.NetworkRetries) retry.Add((entries[i].Id, entries[i].Attempt + 1));
                else exhausted.Add(entries[i].Id);
            }

            if (adaptiveFailure)
            {
                job.GraphConcurrency = Reduce(job.GraphConcurrency, AppConstants.GraphMinConcurrency, AppConstants.GraphMaxConcurrency);
                job.GraphRateLimitRetries += retry.Count;
                if (exhausted.Count > 0)
                {
                    job.Status = JobStatus.Paused;
                    job.StopReason = "graph-temporary";
                    job.LastError = "Microsoft Graph is temporarily rate-limiting or unavailable. No new Outlook changes were made.";
                    _sink.Event("warn", job.LastError + " Processing is paused and can be continued safely after a short wait.");
                    break;
                }
                if (retry.Count > 0)
                {
                    _sink.Event("warn", $"Microsoft Graph temporarily limited {retry.Count} message read{(retry.Count == 1 ? string.Empty : "s")}. Retrying once with concurrency {job.GraphConcurrency}.");
                    JobStore.Save(job);
                    await Task.Delay(Math.Min(AppConstants.MaxGraphRetryDelayMs, retryAfterMs), ct);
                    queue.InsertRange(0, retry);
                }
            }
            else
            {
                job.GraphConcurrency = Grow(job.GraphConcurrency, AppConstants.GraphConcurrencyGrowth, AppConstants.GraphMinConcurrency, AppConstants.GraphMaxConcurrency);
            }
        }
        return output;
    }

    // =====================================================================
    // Jev waves
    // =====================================================================

    private WaveRequest CreateWaveRequest(PendingItem item, MessageMetadata? metadata, IReadOnlyList<LabelRule> rules, string stage, string body)
    {
        if (metadata is null) throw new InvalidOperationException("Outlook metadata is unavailable for a pending message.");
        var payloadText = JevPayloadBuilder.Serialize(JevPayloadBuilder.Build(metadata, rules, stage, body, _model));
        return new WaveRequest
        {
            Item = item,
            Metadata = metadata,
            Stage = stage,
            PayloadText = payloadText,
            EstimatedCost = JevPayloadBuilder.EstimateCost(payloadText),
        };
    }

    private static void AccountCost(TriageJob job, WaveRequest request, JevResult? parsed)
    {
        var reserved = Math.Max(0, request.EstimatedCost);
        var cost = Math.Max(0, parsed?.CostUsd ?? 0);
        job.SpentUsd = Math.Max(0, job.SpentUsd + cost - reserved);
        switch (parsed?.CostSource)
        {
            case "reported": job.ReportedCostUsd += cost; break;
            case "input_tokens": job.TokenCalculatedCostUsd += cost; break;
            default: job.EstimatedCostUsd += cost; break;
        }
    }

    /// <summary>
    /// Dispatch independent Jev requests concurrently with one bounded retry wave.
    /// Requests that do not fit the remaining budget are left unresolved.
    /// Returns true when the budget blocked at least one request.
    /// </summary>
    private async Task<bool> DispatchWaveAsync(List<WaveRequest> requests, IReadOnlyList<LabelRule> rules, TriageJob job, CancellationToken ct)
    {
        var budgetBlocked = false;

        async Task<List<WaveRequest>> Dispatch(List<WaveRequest> candidates, bool isRetry)
        {
            var selected = new List<WaveRequest>();
            double reserve = 0;
            foreach (var request in candidates)
            {
                var next = Math.Max(0, request.EstimatedCost);
                if (job.SpentUsd + reserve + next > job.MaxSpendUsd) { budgetBlocked = true; break; }
                selected.Add(request);
                reserve += next;
            }
            if (selected.Count == 0) return selected;

            job.SpentUsd += reserve;
            job.ModelRequests += selected.Count;
            if (isRetry) job.ProviderRetries += selected.Count;
            JobStore.Save(job);

            var tasks = selected.Select(r => _jev.DecideAsync(r.PayloadText, r.EstimatedCost, rules, ct)).ToList();
            try
            {
                await Task.WhenAll(tasks);
            }
            finally
            {
                for (var i = 0; i < selected.Count; i++)
                {
                    if (tasks[i].IsCompletedSuccessfully)
                    {
                        selected[i].Result = tasks[i].Result;
                        AccountCost(job, selected[i], selected[i].Result);
                    }
                    else
                    {
                        // Cancelled or faulted before an answer: release the reservation; the request
                        // is re-dispatched on resume.
                        job.SpentUsd = Math.Max(0, job.SpentUsd - Math.Max(0, selected[i].EstimatedCost));
                        job.ModelRequests = Math.Max(0, job.ModelRequests - 1);
                    }
                }
                JobStore.Save(job);
            }
            return selected;
        }

        async Task<bool> RetryInAdaptiveChunks(List<WaveRequest> candidates, string message)
        {
            if (candidates.Count == 0) return true;
            var retryAfterMs = AppConstants.JevRetryDelayMs;
            foreach (var request in candidates) retryAfterMs = Math.Max(retryAfterMs, request.Result?.RetryAfterMs ?? 0);
            _sink.Event("warn", message);
            await Task.Delay(Math.Min(AppConstants.MaxJevRetryDelayMs, retryAfterMs), ct);

            for (var offset = 0; offset < candidates.Count;)
            {
                var size = Math.Min(JevWindow(job), candidates.Count - offset);
                var chunk = candidates.GetRange(offset, size);
                var previous = chunk.ToDictionary(r => r, r => r.Result);
                foreach (var request in chunk) request.Result = null;
                var sent = await Dispatch(chunk, true);
                foreach (var request in chunk)
                {
                    if (!sent.Contains(request)) request.Result = previous[request]; // not re-sent: keep the original failure for diagnostics
                }
                if (sent.Count < chunk.Count) return false;
                offset += sent.Count;
                if (sent.Any(r => r.Result is { Ok: false, FailureScope: "session" })) return false;
            }
            return true;
        }

        var attempted = new List<WaveRequest>();
        var cursor = 0;
        while (cursor < requests.Count)
        {
            var size = Math.Min(JevWindow(job), requests.Count - cursor);
            var candidates = requests.GetRange(cursor, size);
            var sent = await Dispatch(candidates, false);
            attempted.AddRange(sent);
            cursor += sent.Count;
            if (sent.Count < candidates.Count) break;

            var sessionFailures = sent.Where(r => r.Result is { Ok: false, FailureScope: "session" }).ToList();
            if (sessionFailures.Count == 0)
            {
                job.JevConcurrency = Grow(job.JevConcurrency, AppConstants.JevConcurrencyGrowth, AppConstants.JevMinConcurrency, AppConstants.JevMaxConcurrency);
                continue;
            }

            var retryableSession = sessionFailures.Where(r => r.Result!.Retryable).ToList();
            if (retryableSession.Count == 0) break;

            job.JevConcurrency = Reduce(job.JevConcurrency, AppConstants.JevMinConcurrency, AppConstants.JevMaxConcurrency);
            var allFailedTogether = retryableSession.Count == sent.Count &&
                                    retryableSession.All(r => r.Result!.Code == retryableSession[0].Result!.Code);

            if (allFailedTogether)
            {
                // Probe one request before retrying the remainder to avoid duplicating a
                // whole paid wave during a persistent provider outage.
                var probe = retryableSession.Take(1).ToList();
                if (!await RetryInAdaptiveChunks(probe, $"Jev is temporarily unavailable. Retrying one request with reduced concurrency {job.JevConcurrency}.")) break;
                if (probe[0].Result is not { Ok: true }) break;
                if (!await RetryInAdaptiveChunks(retryableSession.Skip(1).ToList(),
                        $"The Jev probe succeeded. Retrying the remaining {retryableSession.Count - 1} request(s) once.")) break;
            }
            else if (!await RetryInAdaptiveChunks(retryableSession,
                         $"Jev temporarily limited {retryableSession.Count} request(s). Retrying once with concurrency {job.JevConcurrency}."))
            {
                break;
            }

            if (retryableSession.Any(r => r.Result is not { Ok: true })) break;
        }

        // Contract-validation errors are message-scoped. Retry only as many as the
        // circuit breaker can safely consume.
        var invalid = attempted.Where(r => r.Result is { Ok: false, Retryable: true } && r.Result.FailureScope != "session").ToList();
        invalid = invalid.Take(Math.Max(1, AppConstants.MaxConsecutiveJevFailures - job.ConsecutiveJevFailures)).ToList();
        if (invalid.Count > 0)
        {
            await RetryInAdaptiveChunks(invalid,
                $"Jev response validation failed ({invalid[0].Result!.Code}) for {invalid.Count} response{(invalid.Count == 1 ? string.Empty : "s")}. Retrying once.");
        }

        return budgetBlocked;
    }

    private static int JevWindow(TriageJob job) => NormalizeConcurrency(job.JevConcurrency, AppConstants.JevMinConcurrency, AppConstants.JevMaxConcurrency);

    private static int NormalizeConcurrency(int value, int min, int max) => Math.Clamp(value <= 0 ? min : value, min, max);
    private static int Grow(int value, int growth, int min, int max) => Math.Min(max, NormalizeConcurrency(value, min, max) + growth);
    private static int Reduce(int value, int min, int max) => Math.Max(min, NormalizeConcurrency(value, min, max) / 2);

    // =====================================================================
    // Failure policies (fail closed: never modify Outlook on doubt)
    // =====================================================================

    private void HandleJevFailure(TriageJob job, PendingItem item, MessageMetadata metadata, string stage, JevResult parsed, List<ResultRow> results)
    {
        if (parsed.FailureScope == "session")
        {
            if (!parsed.Retryable) throw new InvalidOperationException(parsed.Error);
            job.Status = JobStatus.Paused;
            job.StopReason = "provider-temporary";
            job.LastError = parsed.Error;
            _sink.Event("warn", parsed.Error + " Processing is paused; run 'jevoutlook continue' to try again later.");
            return;
        }

        job.ConsecutiveJevFailures += 1;
        if (job.ConsecutiveJevFailures >= AppConstants.MaxConsecutiveJevFailures)
        {
            job.Status = JobStatus.Paused;
            job.StopReason = "jev-response-circuit-breaker";
            job.LastError = $"Jev returned invalid responses for {job.ConsecutiveJevFailures} consecutive messages. Processing was paused to protect the remaining API budget. Last issue: {parsed.Error}";
            _sink.Event("warn", job.LastError + " No Outlook changes were made to the current message. Continue later when the provider is stable.");
            return;
        }

        RememberSkipped(job, item.Id);
        job.Skipped += 1;
        job.ModelResponseSkips += 1;
        RemovePending(job, item.Id);
        _sink.Event("warn",
            $"Skipped “{Clip(metadata.Subject, "(no subject)")}” from {Clip(metadata.From, "(unknown sender)")} because the Jev response could not be validated after one retry ({parsed.Code}). {parsed.Error}");
        results.Add(new ResultRow(metadata.From, metadata.Subject, "Not assigned", string.Empty, stage, "skipped-invalid-model-response"));
        job.Target = Math.Max(job.Target, job.HandledCount + job.Pending.Count);
    }

    private void SkipUnreadable(TriageJob job, PendingItem item, MessageMetadata metadata, string reason, List<ResultRow> results)
    {
        RememberSkipped(job, item.Id);
        job.Skipped += 1;
        RemovePending(job, item.Id);
        _sink.Event("warn",
            $"Skipped “{Clip(metadata.Subject, "(no subject)")}” from {Clip(metadata.From, "(unknown sender)")} because the app could not extract readable text. {reason} No safe review fallback is configured, so no categories or archive actions were applied.");
        results.Add(new ResultRow(metadata.From, metadata.Subject, "Not assigned", string.Empty, "metadata", "skipped-no-content"));
        job.Target = Math.Max(job.Target, job.HandledCount + job.Pending.Count);
    }

    private void SkipUnavailable(TriageJob job, string id, MessageMetadata? metadata, List<ResultRow> results)
    {
        RememberSkipped(job, id);
        job.Skipped += 1;
        RemovePending(job, id);
        var safe = metadata ?? MessageMetadata.Empty(id);
        _sink.Event("warn", "A selected Outlook message is no longer available. It was skipped without applying categories or archive actions.");
        results.Add(new ResultRow(safe.From, safe.Subject, "Not assigned", string.Empty, "metadata", "skipped-message-unavailable"));
        job.Target = Math.Max(job.Target, job.HandledCount + job.Pending.Count);
    }

    private static void RememberSkipped(TriageJob job, string id)
    {
        if (job.SkippedMessageIds.Contains(id)) return;
        if (job.SkippedMessageIds.Count >= AppConstants.MaxSkippedMessageIds)
        {
            throw new InvalidOperationException("Too many messages were skipped safely in this session. Start a new session with a narrower Outlook scope.");
        }
        job.SkippedMessageIds.Add(id);
    }

    private static void RemovePending(TriageJob job, string id) => job.Pending.RemoveAll(p => p.Id == id);

    // =====================================================================
    // Outlook writes
    // =====================================================================

    /// <summary>
    /// Categories first (for every message), then archive moves. Both are idempotent:
    /// re-adding a category is a no-op and a message already moved answers 404.
    /// </summary>
    private async Task<WriteOutcome> ApplyFinalResultsAsync(List<FinalResult> finalized, Dictionary<string, MessageMetadata> metadataById,
        Dictionary<string, string> categoryNames, CancellationToken ct)
    {
        // Re-read the current category lists right before writing: the snapshot taken at the start of
        // the batch may be minutes old and a PATCH replaces the whole collection.
        var ids = finalized.Select(f => f.Item.Id).ToList();
        var fresh = await _graph.ReadMessagesAsync(ids, ReadMode.Categories, ct);
        var currentCategories = new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal);
        var gone = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 0; i < ids.Count; i++)
        {
            var part = fresh[i];
            if (part.Ok) { currentCategories[ids[i]] = MessageMetadata.ReadCategories(part.Body); continue; }
            if (part.Status == 404) { gone.Add(ids[i]); continue; } // moved/deleted meanwhile: nothing left to write
            if (part.AuthFailure) return new WriteOutcome(false, false, "Microsoft Graph authorization failed before writing categories. Sign in again (jevoutlook auth).");
            if (!part.Retryable) return new WriteOutcome(false, false, $"Microsoft Graph rejected a category read with HTTP {part.Status}. {part.Error}".Trim());
            return new WriteOutcome(false, true, "Microsoft Graph is temporarily rate-limiting or unavailable while re-reading categories.");
        }

        var updates = new List<(string, IReadOnlyList<string>)>();
        foreach (var final in finalized)
        {
            if (gone.Contains(final.Item.Id)) continue;
            var existing = currentCategories.TryGetValue(final.Item.Id, out var current) ? current
                : metadataById.TryGetValue(final.Item.Id, out var metadata) ? metadata.Categories : [];
            var categoryName = categoryNames.GetValueOrDefault(final.Rule.Name.ToLowerInvariant(), final.Rule.Name);
            var markerName = categoryNames.GetValueOrDefault(AppConstants.TechnicalTriagedCategory.ToLowerInvariant(), AppConstants.TechnicalTriagedCategory);
            var merged = new List<string>(existing);
            foreach (var wanted in new[] { categoryName, markerName })
            {
                if (!merged.Any(c => string.Equals(c, wanted, StringComparison.OrdinalIgnoreCase))) merged.Add(wanted);
            }
            updates.Add((final.Item.Id, merged));
        }

        var outcome = await _graph.PatchCategoriesAsync(updates, ct);
        if (!outcome.Ok) return outcome;

        var archiveIds = finalized.Where(f => f.Archive && !gone.Contains(f.Item.Id)).Select(f => f.Item.Id).ToList();
        return archiveIds.Count > 0 ? await _graph.MoveToArchiveAsync(archiveIds, ct) : WriteOutcome.Success;
    }

    private static string Clip(string? value, string fallback, int max = 160)
    {
        var text = (value ?? string.Empty).Trim();
        if (text.Length == 0) return fallback;
        return text.Length > max ? text[..max] : text;
    }
}
