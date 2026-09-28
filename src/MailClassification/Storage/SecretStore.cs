using System.Diagnostics;

namespace JevOutlook.Storage;

/// <summary>
/// Mailbox passwords. On macOS they live in the login Keychain (service
/// <c>jevoutlook</c>, account = the account id) through the <c>security</c>
/// tool, so nothing secret is written next to the configuration. Elsewhere a
/// <c>secrets.json</c> file with mode 0600 is used.
/// </summary>
public static class SecretStore
{
    private const string Service = "jevoutlook";

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
            var (code, output) = Run(["find-generic-password", "-s", Service, "-a", accountId, "-w"], throwOnError: false);
            return code == 0 ? output.TrimEnd('\n', '\r') : null;
        }
        return LoadFile().GetValueOrDefault(accountId);
    }

    public static string Require(string accountId) =>
        Get(accountId) ?? throw new InvalidOperationException(
            $"No password is stored for account '{accountId}'. Set one with: jevoutlook account password {accountId}");

    public static void Delete(string accountId)
    {
        if (OperatingSystem.IsMacOS())
        {
            Run(["delete-generic-password", "-s", Service, "-a", accountId], throwOnError: false);
            return;
        }
        var all = LoadFile();
        if (all.Remove(accountId)) JsonStore.Save(AppPaths.Secrets, all);
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
