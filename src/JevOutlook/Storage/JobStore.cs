using JevOutlook.Triage;

namespace JevOutlook.Storage;

/// <summary>Checkpointed processing session (one at a time per user).</summary>
public static class JobStore
{
    public static TriageJob? Load() => JsonStore.Load<TriageJob>(AppPaths.Job);

    public static void Save(TriageJob job)
    {
        job.UpdatedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        JsonStore.Save(AppPaths.Job, job);
    }

    public static void Delete()
    {
        JsonStore.Delete(AppPaths.Job);
        RuleStore.DeleteJobRules();
    }
}
