namespace JevOutlook.Storage;

/// <summary>
/// Everything jevOutlook persists lives in one user directory (default
/// <c>~/.jevoutlook</c>, overridable with <c>JEVOUTLOOK_HOME</c>). Global files
/// (config, rules, accounts) sit at the root; per-account state (sign-in record,
/// processing session) lives under <c>accounts/&lt;id&gt;/</c>.
/// </summary>
public static class AppPaths
{
    public static string Root
    {
        get
        {
            var custom = Environment.GetEnvironmentVariable("JEVOUTLOOK_HOME");
            if (!string.IsNullOrWhiteSpace(custom)) return custom;
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(home, ".jevoutlook");
        }
    }

    public static string Config => Path.Combine(Root, "config.json");
    public static string Rules => Path.Combine(Root, "rules.json");
    public static string Accounts => Path.Combine(Root, "accounts.json");
    public static string Secrets => Path.Combine(Root, "secrets.json");
    public static string TokenCacheName => "jevoutlook";

    public static string AccountsRoot => Path.Combine(Root, "accounts");
    public static string AccountDir(string accountId) => Path.Combine(AccountsRoot, accountId);
    public static string AccountAuthRecord(string accountId) => Path.Combine(AccountDir(accountId), "auth-record.json");
    public static string AccountJob(string accountId) => Path.Combine(AccountDir(accountId), "job.json");
    public static string AccountJobRules(string accountId) => Path.Combine(AccountDir(accountId), "job-rules.json");
    public static string AccountJobLock(string accountId) => Path.Combine(AccountDir(accountId), "job.lock");

    // Pre-multi-account layout (single Microsoft mailbox), migrated on first start.
    public static string LegacyJob => Path.Combine(Root, "job.json");
    public static string LegacyJobRules => Path.Combine(Root, "job-rules.json");
    public static string LegacyAuthRecord => Path.Combine(Root, "auth-record.json");
}
