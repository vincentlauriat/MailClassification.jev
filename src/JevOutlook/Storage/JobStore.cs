using JevOutlook.Rules;
using JevOutlook.Triage;

namespace JevOutlook.Storage;

/// <summary>Checkpointed processing session of one account (one session at a time per mailbox).</summary>
public sealed class JobStore
{
    private readonly string _jobPath;
    private readonly string _rulesPath;

    public JobStore(string accountId)
    {
        AccountId = accountId;
        _jobPath = AppPaths.AccountJob(accountId);
        _rulesPath = AppPaths.AccountJobRules(accountId);
    }

    public string AccountId { get; }

    public TriageJob? Load() => JsonStore.Load<TriageJob>(_jobPath);

    public void Save(TriageJob job)
    {
        job.UpdatedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        JsonStore.Save(_jobPath, job);
    }

    public void Delete()
    {
        JsonStore.Delete(_jobPath);
        JsonStore.Delete(_rulesPath);
    }

    /// <summary>Rules frozen for the current job so later edits cannot change a running classification.</summary>
    public void SaveRules(IEnumerable<LabelRule> rules) => JsonStore.Save(_rulesPath, rules.ToList());

    public List<LabelRule> LoadRules()
    {
        var rules = JsonStore.Load<List<LabelRule>>(_rulesPath)
            ?? throw new InvalidOperationException("The rules of the current processing session are missing. Clear the session and start again.");
        return RuleValidator.ValidateAndNormalize(rules);
    }
}
