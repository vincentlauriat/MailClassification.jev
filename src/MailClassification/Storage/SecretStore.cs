using System.Diagnostics;

namespace MailClassification.Storage;

/// <summary>
/// Mailbox passwords. On macOS they live in the login Keychain (service
/// <c>mailclassification</c>, account = the account id) through the <c>security</c>
/// tool, so nothing secret is written next to the configuration. Elsewhere a
/// <c>secrets.json</c> file with mode 0600 is used (it moves with the state directory).
/// Items stored under the pre-rename service <c>jevoutlook</c> are migrated on first read.
/// </summary>
public static class SecretStore
{
    private const string Service = "mailclassification";
    private const string LegacyService = "jevoutlook";

    public static string Backend => OperatingSystem.IsMacOS() ? "macOS Keychain" : AppPaths.Secrets;

    public static void Set(string accountId, string secret)
    {
        if (OperatingSystem.IsMacOS())
        {
            // -U updates an existing item in place; the password is passed as an argument to the
            // child process only (never through the shell), which is how Apple's own tooling does it.
            Run(["add-generic-password", "-U", "-s", Service, "-a", accountId, "-l", $"{Service}: {accountId}", "-w", secret]);
            return;
        }
        var all = LoadFile();
        all[accountId] = secret;
        JsonStore.Save(AppPaths.Secrets, all);
    }

    public static string? Get(string accountId)
    {
        if (OperatingSystem.IsMacOS())
        {
            return GetWithLegacyFallback(
                service => Find(service, accountId),
                secret => Set(accountId, secret),
                () => Run(["delete-generic-password", "-s", LegacyService, "-a", accountId], throwOnError: false));
        }
        return LoadFile().GetValueOrDefault(accountId);
    }

    public static string Require(string accountId) =>
        Get(accountId) ?? throw new InvalidOperationException(
            $"No password is stored for account '{accountId}'. Set one with: mailclassification account password {accountId}");

    public static void Delete(string accountId)
    {
        if (OperatingSystem.IsMacOS())
        {
            Run(["delete-generic-password", "-s", Service, "-a", accountId], throwOnError: false);
            Run(["delete-generic-password", "-s", LegacyService, "-a", accountId], throwOnError: false); // never migrated
            return;
        }
        var all = LoadFile();
        if (all.Remove(accountId)) JsonStore.Save(AppPaths.Secrets, all);
    }

    /// <summary>
    /// Read under the current service; on a miss, read the legacy item and, when found, store it
    /// under the current service and delete the legacy one. The legacy item is only deleted
    /// after the new one was written; if the write fails the legacy value is still returned
    /// and the migration is retried on the next read.
    /// </summary>
    internal static string? GetWithLegacyFallback(Func<string, string?> find, Action<string> store, Action deleteLegacy)
    {
        if (find(Service) is { } current) return current;
        if (find(LegacyService) is not { } legacy) return null;
        try
        {
            store(legacy);
        }
        catch (InvalidOperationException)
        {
            return legacy;
        }
        deleteLegacy();
        return legacy;
    }

    private static string? Find(string service, string accountId)
    {
        var (code, output) = Run(["find-generic-password", "-s", service, "-a", accountId, "-w"], throwOnError: false);
        return code == 0 ? output.TrimEnd('\n', '\r') : null;
    }

    public static bool Has(string accountId) => Get(accountId) is { Length: > 0 };

    private static Dictionary<string, string> LoadFile() =>
        JsonStore.Load<Dictionary<string, string>>(AppPaths.Secrets) ?? new Dictionary<string, string>(StringComparer.Ordinal);

    private static (int Code, string Output) Run(string[] args, bool throwOnError = true)
    {
        var psi = new ProcessStartInfo("/usr/bin/security")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        foreach (var a in args) psi.ArgumentList.Add(a);
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Could not start /usr/bin/security.");
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (throwOnError && process.ExitCode != 0)
            throw new InvalidOperationException("The macOS Keychain refused the operation: " + (error.Trim().Length > 0 ? error.Trim() : $"exit code {process.ExitCode}"));
        return (process.ExitCode, output);
    }
}
