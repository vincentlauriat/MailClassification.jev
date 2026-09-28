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

    /// <summary>Shown when another process (dashboard or CLI) holds this account's job lock.</summary>
    public const string LockedMessage =
        "Another jevOutlook process is processing this mailbox right now (dashboard or CLI). Wait for it or stop it, then retry.";

    /// <summary>
    /// Exclusive, cross-process lock on this account's session (<c>accounts/&lt;id&gt;/job.lock</c>),
    /// or null when another holder has it. Without it the dashboard (always running as a
    /// LaunchAgent) and <c>jevoutlook continue</c> could process the same session at once: both pay
    /// for the same messages and each overwrites the other's spend and status. Non-blocking.
    /// </summary>
    public IDisposable? TryLock() => TryLockFile(AppPaths.AccountJobLock(AccountId));

    /// <summary><see cref="TryLock"/>, or an <see cref="InvalidOperationException"/> carrying <see cref="LockedMessage"/>.</summary>
    public IDisposable Lock() => TryLock() ?? throw new InvalidOperationException(LockedMessage);

    /// <summary>
    /// <see cref="FileShare.None"/> is an advisory <c>flock(LOCK_EX | LOCK_NB)</c> on Unix and a
    /// share-mode lock on Windows; a second open fails even within the same process. The file is
    /// never deleted: unlinking it while held would let a second process lock a new inode.
    /// </summary>
    internal static IDisposable? TryLockFile(string path)
    {
        JsonStore.EnsureDirectory(Path.GetDirectoryName(path)!);
        try
        {
            var options = new FileStreamOptions { Mode = FileMode.OpenOrCreate, Access = FileAccess.ReadWrite, Share = FileShare.None };
            if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            var stream = new FileStream(path, options);
            // A lock file created by an earlier build may be 0644: every state file is 0600.
            if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
            return stream;
        }
        catch (IOException)
        {
            return null;
        }
    }

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
