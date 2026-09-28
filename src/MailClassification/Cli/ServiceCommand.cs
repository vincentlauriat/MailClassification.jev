using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using MailClassification.Storage;
using MailClassification.Web;

namespace MailClassification.Cli;

/// <summary>
/// <c>mailclassification service</c>: a macOS LaunchAgent that keeps <c>ui --no-open</c> running
/// from login onward, so the Outlook add-in pane (served by that process) always loads.
/// launchctl is always called with an argument list, never through a shell.
/// </summary>
internal static class ServiceCommand
{
    public const string Label = "com.vincentlauriat.mailclassification";

    /// <summary>Label of the agent installed by jevOutlook builds; replaced by <c>service install</c>.</summary>
    public const string LegacyLabel = "com.vincentlauriat.jevoutlook";

    /// <summary>
    /// The only variables copied from the installing shell into the plist. API keys
    /// (OPENROUTER_API_KEY, JEV_API_KEY) never go there: the plist is plaintext.
    /// </summary>
    public static readonly IReadOnlyList<string> AllowedEnvironment = ["DOTNET_ROOT", AppPaths.HomeVariable, AppPaths.LegacyHomeVariable];

    public static string LaunchAgentsDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "LaunchAgents");

    public static string PlistPath => Path.Combine(LaunchAgentsDir, Label + ".plist");

    /// <summary>The jevOutlook-era plist in <paramref name="launchAgentsDir"/>, or null when there is none.</summary>
    internal static string? LegacyPlistIn(string launchAgentsDir)
    {
        var path = Path.Combine(launchAgentsDir, LegacyLabel + ".plist");
        return File.Exists(path) ? path : null;
    }

    public static string LogPath => Path.Combine(AppPaths.Root, "logs", "ui.log");

    public static async Task<int> RunAsync(string[] args, HttpClient http, CancellationToken ct)
    {
        if (!OperatingSystem.IsMacOS())
            throw new PlatformNotSupportedException("'mailclassification service' manages a macOS LaunchAgent. On other systems, run 'mailclassification ui --no-open' from your own service manager.");
        var sub = args.Length > 0 ? args[0].ToLowerInvariant() : "status";
        return sub switch
        {
            "install" => await InstallAsync(args, http, ct),
            "uninstall" or "remove" => Uninstall(),
            "status" => await StatusAsync(args, http, ct),
            "restart" => Restart(),
            _ => throw new ArgumentException("Usage: mailclassification service install [--port 5177] [--https-port 5178] [--exe <path>] | status | restart | uninstall"),
        };
    }

    // ---------------------------------------------------------------------
    // subcommands
    // ---------------------------------------------------------------------

    private static async Task<int> InstallAsync(string[] args, HttpClient http, CancellationToken ct)
    {
        var port = (int)Program.Number(args, "--port", 5177);
        var httpsPort = (int)Program.Number(args, "--https-port", 5178);
        if (port is < 1 or > 65535 || httpsPort is < 1 or > 65535) throw new ArgumentException("Ports must be between 1 and 65535.");

        var exe = Path.GetFullPath(Program.Option(args, "--exe") ?? Environment.ProcessPath
            ?? throw new InvalidOperationException("Cannot determine the mailclassification executable; pass --exe <path>."));
        if (Path.GetFileNameWithoutExtension(exe).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Running through the 'dotnet' host: pass the mailclassification executable with --exe <path>.");
        if (!File.Exists(exe)) throw new FileNotFoundException($"Executable not found: {exe}");
        if (exe.Contains("/bin/Debug/", StringComparison.Ordinal) || exe.Contains("/bin/Release/", StringComparison.Ordinal))
        {
            Console.Error.WriteLine($"Warning: {exe} is a build output; the next build or 'dotnet clean' will replace or remove it under the agent. Prefer a published copy:");
            Console.Error.WriteLine("  dotnet publish src/MailClassification -c Release -o ~/.mailclassification/bin");
            Console.Error.WriteLine("  ~/.mailclassification/bin/mailclassification service install --exe ~/.mailclassification/bin/mailclassification");
        }

        var env = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var name in AllowedEnvironment)
        {
            if (Environment.GetEnvironmentVariable(name) is { Length: > 0 } value) env[name] = value;
        }
        // launchd starts the agent without the shell's environment: the framework-dependent
        // apphost then only finds .NET through DOTNET_ROOT or the default install locations.
        if (!env.ContainsKey("DOTNET_ROOT") && !File.Exists("/usr/local/share/dotnet/dotnet") && !File.Exists("/etc/dotnet/install_location"))
            Console.Error.WriteLine("Warning: DOTNET_ROOT is not set and .NET is not in /usr/local/share/dotnet; the agent may not find the runtime. Export DOTNET_ROOT and install again.");
        if (Environment.GetEnvironmentVariable("OPENROUTER_API_KEY") is { Length: > 0 } || Environment.GetEnvironmentVariable("JEV_API_KEY") is { Length: > 0 })
            Console.Error.WriteLine("Note: the API key from your environment is NOT copied into the agent (the plist is plaintext). Store it with: mailclassification key set <api-key>");

        Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
        Directory.CreateDirectory(Path.GetDirectoryName(PlistPath)!);
        string[] programArgs = ["ui", "--no-open", "--port", port.ToString(CultureInfo.InvariantCulture), "--https-port", httpsPort.ToString(CultureInfo.InvariantCulture)];
        File.WriteAllText(PlistPath, BuildLaunchAgentPlist(exe, programArgs, env, LogPath), new UTF8Encoding(false));

        var domain = Domain();
        // The jevOutlook agent holds the same port: stop and remove it before bootstrapping, and
        // before the /health probe below (which also accepts the legacy server's answer).
        var legacyLoaded = IsLoaded(domain, LegacyLabel);
        var legacyPlist = LegacyPlistIn(LaunchAgentsDir);
        if (legacyLoaded) Launchctl("bootout", $"{domain}/{LegacyLabel}");
        if (legacyPlist is not null) File.Delete(legacyPlist);
        if (legacyLoaded || legacyPlist is not null)
            Console.WriteLine($"Removed the previous jevOutlook LaunchAgent ({LegacyLabel}).");
        if (IsLoaded(domain, Label)) Launchctl("bootout", $"{domain}/{Label}");
        // bootout returns before launchd has fully released the job; retry briefly.
        var (code, output) = (0, string.Empty);
        for (var attempt = 0; attempt < 5; attempt++)
        {
            (code, output) = Launchctl("bootstrap", domain, PlistPath);
            if (code == 0) break;
            await Task.Delay(TimeSpan.FromSeconds(1), ct);
        }
        if (code != 0) throw new InvalidOperationException($"launchctl bootstrap failed ({code}): {output.Trim()}");

        Console.WriteLine($"LaunchAgent installed: {PlistPath}");
        Console.WriteLine($"  runs  : {exe} {string.Join(' ', programArgs)}");
        Console.WriteLine($"  log   : {LogPath}");
        var state = UiServer.HealthState.Free;
        for (var i = 0; i < 20 && state != UiServer.HealthState.Ours; i++)
        {
            await Task.Delay(TimeSpan.FromMilliseconds(500), ct);
            state = (await UiServer.ProbeAsync(http, port, ct)).State;
        }
        if (state == UiServer.HealthState.Ours)
            Console.WriteLine($"MailClassification is up: http://127.0.0.1:{port}/ · add-in pane https://localhost:{httpsPort}/taskpane.html");
        else if (state == UiServer.HealthState.Foreign)
            Console.WriteLine($"Port {port} is held by another program or an older build without /health: stop it; the agent retries every 30 s. See {LogPath}.");
        else
            Console.WriteLine($"The agent is loaded but /health does not answer yet. Check 'mailclassification service status' and {LogPath}.");
        return 0;
    }

    private static int Uninstall()
    {
        var domain = Domain();
        var loaded = IsLoaded(domain, Label);
        if (loaded)
        {
            var (code, output) = Launchctl("bootout", $"{domain}/{Label}");
            if (code != 0) throw new InvalidOperationException($"launchctl bootout failed ({code}): {output.Trim()}");
        }
        var existed = File.Exists(PlistPath);
        if (existed) File.Delete(PlistPath);
        Console.WriteLine(loaded || existed ? "LaunchAgent stopped and removed." : "No LaunchAgent was installed.");
        ReportLegacyAgent(domain);
        return 0;
    }

    private static async Task<int> StatusAsync(string[] args, HttpClient http, CancellationToken ct)
    {
        var port = (int)Program.Number(args, "--port", 5177);
        var installed = File.Exists(PlistPath);
        var domain = Domain();
        var loaded = IsLoaded(domain, Label);
        var (state, version) = await UiServer.ProbeAsync(http, port, ct);
        Console.WriteLine($"Plist   : {(installed ? PlistPath : "not installed (mailclassification service install)")}");
        Console.WriteLine($"Agent   : {(loaded ? "loaded" : "not loaded")} ({Label})");
        Console.WriteLine($"Health  : {state switch
        {
            UiServer.HealthState.Ours => $"MailClassification {version} answers on http://127.0.0.1:{port}/",
            UiServer.HealthState.Foreign => $"port {port} is held by another program (or an older build without /health)",
            _ => $"nothing listens on port {port}",
        }}");
        Console.WriteLine($"Log     : {LogPath}");
        ReportLegacyAgent(domain);
        return state == UiServer.HealthState.Ours ? 0 : 1;
    }

    private static int Restart()
    {
        var (code, output) = Launchctl("kickstart", "-k", $"{Domain()}/{Label}");
        if (code != 0) throw new InvalidOperationException($"launchctl kickstart failed ({code}): {output.Trim()} — is the agent installed? Run 'mailclassification service install'.");
        Console.WriteLine("LaunchAgent restarted.");
        return 0;
    }

    // ---------------------------------------------------------------------
    // plist
    // ---------------------------------------------------------------------

    /// <summary>
    /// The LaunchAgent property list. Pure (no I/O) so it can be unit-tested; values are
    /// XML-escaped by <see cref="XElement"/>. Only <see cref="AllowedEnvironment"/> entries of
    /// <paramref name="env"/> are written.
    /// </summary>
    public static string BuildLaunchAgentPlist(string exe, IReadOnlyList<string> args, IReadOnlyDictionary<string, string> env, string logPath)
    {
        var dict = new XElement("dict",
            new XElement("key", "Label"), new XElement("string", Label),
            new XElement("key", "ProgramArguments"),
            new XElement("array", new[] { exe }.Concat(args).Select(a => new XElement("string", a))),
            new XElement("key", "RunAtLoad"), new XElement("true"),
            new XElement("key", "KeepAlive"), new XElement("true"),
            new XElement("key", "ThrottleInterval"), new XElement("integer", 30),
            new XElement("key", "StandardOutPath"), new XElement("string", logPath),
            new XElement("key", "StandardErrorPath"), new XElement("string", logPath));

        var kept = env.Where(kv => AllowedEnvironment.Contains(kv.Key)).OrderBy(kv => kv.Key, StringComparer.Ordinal).ToList();
        if (kept.Count > 0)
        {
            dict.Add(new XElement("key", "EnvironmentVariables"),
                new XElement("dict", kept.SelectMany(kv => new[] { new XElement("key", kv.Key), new XElement("string", kv.Value) })));
        }

        var doc = new XDocument(
            new XDocumentType("plist", "-//Apple//DTD PLIST 1.0//EN", "http://www.apple.com/DTDs/PropertyList-1.0.dtd", null),
            new XElement("plist", new XAttribute("version", "1.0"), dict));
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + doc + "\n";
    }

    // ---------------------------------------------------------------------
    // launchctl
    // ---------------------------------------------------------------------

    private static string Domain()
    {
        var (code, output) = Run("/usr/bin/id", "-u");
        return code == 0 && int.TryParse(output.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var uid)
            ? $"gui/{uid}"
            : throw new InvalidOperationException("Cannot determine the user id ('id -u' failed).");
    }

    private static bool IsLoaded(string domain, string label) => Launchctl("print", $"{domain}/{label}").Code == 0;

    /// <summary>Mention a leftover jevOutlook agent (it is removed by <c>service install</c>).</summary>
    private static void ReportLegacyAgent(string domain)
    {
        var plist = LegacyPlistIn(LaunchAgentsDir);
        var loaded = IsLoaded(domain, LegacyLabel);
        if (plist is null && !loaded) return;
        Console.WriteLine($"Legacy  : the jevOutlook agent {LegacyLabel} is still {(loaded ? "loaded" : "installed")}{(plist is not null ? $" ({plist})" : string.Empty)}. 'mailclassification service install' replaces it; to only remove it: launchctl bootout {domain}/{LegacyLabel}, then delete its plist.");
    }

    private static (int Code, string Output) Launchctl(params string[] args) => Run("/bin/launchctl", args);

    private static (int Code, string Output) Run(string file, params string[] args)
    {
        var psi = new ProcessStartInfo(file)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        foreach (var a in args) psi.ArgumentList.Add(a);
        using var process = Process.Start(psi) ?? throw new InvalidOperationException($"Could not start {file}.");
        var stderr = process.StandardError.ReadToEndAsync();
        var stdout = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, stdout + stderr.Result);
    }
}
