namespace JevOutlook.Storage;

/// <summary>
/// Everything jevOutlook persists lives in one user directory (default
/// <c>~/.jevoutlook</c>, overridable with <c>JEVOUTLOOK_HOME</c>). This is the
/// counterpart of Apps Script User Properties in the Gmail version.
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
    public static string Job => Path.Combine(Root, "job.json");
    public static string JobRules => Path.Combine(Root, "job-rules.json");
    public static string AuthRecord => Path.Combine(Root, "auth-record.json");
    public static string TokenCacheName => "jevoutlook";
}
