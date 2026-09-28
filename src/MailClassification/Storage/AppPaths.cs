namespace MailClassification.Storage;

/// <summary>
/// Everything MailClassification persists lives in one user directory (default
/// <c>~/.mailclassification</c>, overridable with <c>MAILCLASSIFICATION_HOME</c>; the
/// pre-rename <c>JEVOUTLOOK_HOME</c> is still honoured). Global files (config, rules,
/// accounts) sit at the root; per-account state (sign-in record, processing session)
/// lives under <c>accounts/&lt;id&gt;/</c>.
/// </summary>
public static class AppPaths
{
    public const string HomeVariable = "MAILCLASSIFICATION_HOME";
    public const string DebugVariable = "MAILCLASSIFICATION_DEBUG";

    // Names used while the product was called jevOutlook; still read as fallbacks.
    public const string LegacyHomeVariable = "JEVOUTLOOK_HOME";
    public const string LegacyDebugVariable = "JEVOUTLOOK_DEBUG";
    public const string LegacyDirectoryName = ".jevoutlook";

    public const string DirectoryName = ".mailclassification";

    public static string Root => CustomRoot ?? Path.Combine(UserHome, DirectoryName);

    /// <summary>The state directory set through the environment, or null for the default location.</summary>
    public static string? CustomRoot =>
        Environment.GetEnvironmentVariable(HomeVariable) is { } custom && !string.IsNullOrWhiteSpace(custom) ? custom
        : Environment.GetEnvironmentVariable(LegacyHomeVariable) is { } legacy && !string.IsNullOrWhiteSpace(legacy) ? legacy
        : null;

    public static bool DebugEnabled =>
        !string.IsNullOrEmpty(Environment.GetEnvironmentVariable(DebugVariable)) ||
        !string.IsNullOrEmpty(Environment.GetEnvironmentVariable(LegacyDebugVariable));

    private static string UserHome => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    public static string Config => Path.Combine(Root, "config.json");
    public static string Rules => Path.Combine(Root, "rules.json");
    public static string Accounts => Path.Combine(Root, "accounts.json");
    public static string Secrets => Path.Combine(Root, "secrets.json");

    /// <summary>
    /// Name of the MSAL persistent token cache. Deliberately kept at the pre-rename value
    /// "jevoutlook": it is an internal identifier (never shown), and renaming it would drop
    /// every cached Microsoft sign-in and force a new device-code login for each account.
    /// </summary>
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

    /// <summary>
    /// One-time move of the pre-rename <c>~/.jevoutlook</c> into <c>~/.mailclassification</c>.
    /// Skipped when the state directory is set through the environment. Returns a one-line
    /// notice when something was moved, null otherwise.
    /// </summary>
    public static string? MigrateLegacyRoot() =>
        CustomRoot is not null ? null : MigrateRoot(Path.Combine(UserHome, LegacyDirectoryName), Root);

    /// <summary>
    /// Moves <paramref name="legacyRoot"/> into <paramref name="newRoot"/> when the new root has
    /// no state yet. The new root counts as empty when it is missing or holds nothing but
    /// <c>bin/</c> — the recommended install publishes the executable to
    /// <c>~/.mailclassification/bin</c> before it ever runs. In that case the legacy entries are
    /// moved one by one and names already present (the legacy <c>bin/</c>) stay behind.
    /// Idempotent: once state is in the new root, nothing happens.
    /// </summary>
    internal static string? MigrateRoot(string legacyRoot, string newRoot)
    {
        if (!Directory.Exists(legacyRoot)) return null;
        if (!Directory.Exists(newRoot))
        {
            var parent = Path.GetDirectoryName(Path.GetFullPath(newRoot));
            if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
            Directory.Move(legacyRoot, newRoot);
            return $"Moved the jevOutlook state directory {legacyRoot} to {newRoot}.";
        }

        var existing = Directory.EnumerateFileSystemEntries(newRoot).Select(Path.GetFileName).ToList();
        if (existing.Any(name => name != "bin")) return null; // the new root already has state

        var moved = 0;
        foreach (var entry in Directory.EnumerateFileSystemEntries(legacyRoot).ToList())
        {
            var target = Path.Combine(newRoot, Path.GetFileName(entry));
            if (File.Exists(target) || Directory.Exists(target)) continue;
            if (Directory.Exists(entry)) Directory.Move(entry, target);
            else File.Move(entry, target);
            moved++;
        }
        if (moved == 0) return null;
        var left = Directory.EnumerateFileSystemEntries(legacyRoot).Any();
        if (!left) Directory.Delete(legacyRoot);
        return $"Moved the jevOutlook state from {legacyRoot} to {newRoot}." +
               (left ? $" {legacyRoot} now only holds the old executable; delete it once 'mailclassification service install' has replaced the old agent." : string.Empty);
    }
}
