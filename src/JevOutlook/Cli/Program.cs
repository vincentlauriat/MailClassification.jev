using System.Globalization;
using System.Text.Json;
using JevOutlook.Graph;
using JevOutlook.Jev;
using JevOutlook.Rules;
using JevOutlook.Storage;
using JevOutlook.Triage;

namespace JevOutlook.Cli;

public static class Program
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(100) };

    public static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

        try
        {
            var command = args.Length > 0 ? args[0].ToLowerInvariant() : "help";
            var rest = args.Skip(1).ToArray();
            return command switch
            {
                "help" or "--help" or "-h" => Help(),
                "version" or "--version" => Version(),
                "auth" => await AuthAsync(rest, cts.Token),
                "config" => Config(rest),
                "key" => await KeyAsync(rest, cts.Token),
                "playbooks" => Playbooks(),
                "rules" => Rules(rest),
                "run" => await RunAsync(rest, cts.Token),
                "continue" or "resume" => await ContinueAsync(rest, cts.Token),
                "status" => Status(),
                "clear-job" or "clear" => ClearJob(),
                _ => Unknown(command),
            };
        }
        catch (OperationCanceledException)
        {
            Console.Error.WriteLine("Cancelled.");
            return 130;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Error: " + ex.Message);
            if (Environment.GetEnvironmentVariable("JEVOUTLOOK_DEBUG") is { Length: > 0 }) Console.Error.WriteLine(ex);
            return 1;
        }
    }

    // ---------------------------------------------------------------------
    // help / version
    // ---------------------------------------------------------------------

    private static int Help()
    {
        Console.WriteLine("""
            jevoutlook — AI-assisted Outlook classification with Jev (TypeSafe) via OpenRouter

            USAGE
              jevoutlook <command> [options]

            SETUP
              config set client-id <guid>        Entra ID app registration (client) id
              config set tenant-id <id>          common (default) | organizations | consumers | <tenant guid>
              config set provider <name>         openrouter (default) | typesafe
              config set device-code true|false  Use the device-code flow instead of a browser
              config show
              auth [--device-code]               Sign in to Microsoft 365 / Outlook.com
              auth --status | --logout
              key set <api-key>                  Store the OpenRouter (or TypeSafe) API key
              key test [<api-key>]               Send one tiny Jev request to verify the key
              key clear

            CATEGORIES
              playbooks                          List ready-made category sets
              rules show                         Show the saved classification rules
              rules use <playbook-id>            Load a playbook as the saved rules
              rules import <file.json>           Load rules from a JSON file (validated)
              rules export <file.json>           Write the saved rules to a JSON file
              rules reset                        Restore the Universal inbox playbook
              rules path                         Print the rules file location

            PROCESSING
              run [options]                      Start a session (PREVIEW by default)
                --scope inbox|all                Inbox (default) or the whole mailbox
                --include-read                   Also process read messages (default: unread only)
                --limit N|all                    Number of messages (default 10)
                --mode labels|labels-archive     Categories only (default) or archive eligible categories
                --live                           Apply changes to Outlook (default: preview only)
                --max-spend <usd>                Run budget (default 0.10)
                --metadata-threshold <0.5-0.99>  Confidence required to trust the metadata pass (default 0.75)
                --archive-threshold <0.5-0.999>  Confidence required to archive (default 0.93)
                --yes                            Skip the confirmation for live archive runs
              continue [--max-spend <usd>]       Resume a paused session (raise the budget if needed)
              status                             Show the current session
              clear-job                          Forget a finished session

            ENVIRONMENT
              OPENROUTER_API_KEY / JEV_API_KEY   API key (takes precedence over the stored key)
              JEVOUTLOOK_HOME                    State directory (default ~/.jevoutlook)
            """);
        return 0;
    }

    private static int Version()
    {
        Console.WriteLine("jevoutlook " + (typeof(Program).Assembly.GetName().Version?.ToString(3) ?? "0.0.0"));
        return 0;
    }

    private static int Unknown(string command)
    {
        Console.Error.WriteLine($"Unknown command '{command}'. Run 'jevoutlook help'.");
        return 2;
    }

    // ---------------------------------------------------------------------
    // config / auth
    // ---------------------------------------------------------------------

    private static int Config(string[] args)
    {
        var config = AppConfig.Load();
        if (args.Length == 0 || args[0] == "show")
        {
            Console.WriteLine($"State directory : {AppPaths.Root}");
            Console.WriteLine($"Client id       : {config.ClientId ?? "(not set)"}");
            Console.WriteLine($"Tenant id       : {config.TenantId}");
            Console.WriteLine($"Provider        : {config.Provider} → {config.ResolvedEndpoint}");
            Console.WriteLine($"Model           : {config.ResolvedModel}");
            Console.WriteLine($"Device code     : {config.DeviceCode}");
            Console.WriteLine($"Stored API key  : {(string.IsNullOrEmpty(config.ApiKey) ? "no" : "yes")}");
            return 0;
        }
        if (args[0] == "set" && args.Length >= 3)
        {
            var value = string.Join(' ', args.Skip(2)).Trim();
            switch (args[1].ToLowerInvariant())
            {
                case "client-id": config.ClientId = value; break;
                case "tenant-id": config.TenantId = value.Length > 0 ? value : "common"; break;
                case "provider":
                    if (value is not ("openrouter" or "typesafe")) throw new ArgumentException("Provider must be 'openrouter' or 'typesafe'.");
                    config.Provider = value; break;
                case "endpoint": config.EndpointUrl = value.Length > 0 ? value : null; break;
                case "model": config.Model = value.Length > 0 ? value : null; break;
                case "device-code": config.DeviceCode = value is "true" or "1" or "yes"; break;
                default: throw new ArgumentException($"Unknown setting '{args[1]}'.");
            }
            config.Save();
            Console.WriteLine("Saved.");
            return 0;
        }
        throw new ArgumentException("Usage: jevoutlook config show | config set <client-id|tenant-id|provider|endpoint|model|device-code> <value>");
    }

    private static async Task<int> AuthAsync(string[] args, CancellationToken ct)
    {
        var config = AppConfig.Load();
        if (args.Contains("--logout"))
        {
            GraphTokenProvider.DeleteRecord();
            Console.WriteLine("Signed out (saved authentication record removed).");
            return 0;
        }
        if (args.Contains("--status"))
        {
            var record = await GraphTokenProvider.LoadRecordAsync(ct);
            if (record is null) { Console.WriteLine("Not signed in. Run: jevoutlook auth"); return 1; }
            var tokens = new GraphTokenProvider(config, record, config.DeviceCode);
            var graph = new GraphMailClient(Http, tokens);
            Console.WriteLine("Signed in as " + await graph.GetSignedInUserAsync(ct));
            return 0;
        }

        var provider = new GraphTokenProvider(config, null, args.Contains("--device-code"));
        Console.WriteLine("Opening the Microsoft sign-in flow…");
        var newRecord = await provider.SignInAsync(ct);
        await GraphTokenProvider.SaveRecordAsync(newRecord, ct);
        var client = new GraphMailClient(Http, provider);
        Console.WriteLine("Signed in as " + await client.GetSignedInUserAsync(ct));
        return 0;
    }

    private static async Task<GraphMailClient> ConnectGraphAsync(AppConfig config, CancellationToken ct)
    {
        var record = await GraphTokenProvider.LoadRecordAsync(ct)
            ?? throw new InvalidOperationException("Not signed in to Microsoft 365 / Outlook.com. Run: jevoutlook auth");
        return new GraphMailClient(Http, new GraphTokenProvider(config, record, config.DeviceCode));
    }

    // ---------------------------------------------------------------------
    // API key
    // ---------------------------------------------------------------------

    private static async Task<int> KeyAsync(string[] args, CancellationToken ct)
    {
        var config = AppConfig.Load();
        var sub = args.Length > 0 ? args[0] : "test";
        switch (sub)
        {
            case "set":
                if (args.Length < 2 || string.IsNullOrWhiteSpace(args[1])) throw new ArgumentException("Usage: jevoutlook key set <api-key>");
                config.ApiKey = args[1].Trim();
                config.Save();
                Console.WriteLine($"{config.ProviderLabel} API key stored in {AppPaths.Config} (mode 0600).");
                return 0;
            case "clear":
                config.ApiKey = null;
                config.Save();
                Console.WriteLine("Stored API key removed.");
                return 0;
            case "test":
            {
                var key = config.ResolveApiKey(args.Length > 1 ? args[1] : null);
                var rules = RuleStore.Load();
                var payload = JevPayloadBuilder.Serialize(JevPayloadBuilder.Build(JevPayloadBuilder.SampleMetadata(), rules, "metadata", string.Empty, config.ResolvedModel));
                var jev = new JevClient(Http, config.ResolvedEndpoint, key);
                var parsed = await jev.DecideAsync(payload, JevPayloadBuilder.EstimateCost(payload), rules, ct);
                if (!parsed.Ok) throw new InvalidOperationException(parsed.Error);
                Console.WriteLine($"OK — model {(parsed.Model.Length > 0 ? parsed.Model : config.ResolvedModel)}" +
                                  (parsed.Provider.Length > 0 ? $" via {parsed.Provider}" : string.Empty));
                Console.WriteLine($"Sample email → {RuleValidator.NameById(rules, parsed.RuleId)} (confidence {parsed.Confidence.ToString("0.000", CultureInfo.InvariantCulture)}, cost ${parsed.CostUsd.ToString("0.000000", CultureInfo.InvariantCulture)} {parsed.CostSource})");
                return 0;
            }
            default:
                throw new ArgumentException("Usage: jevoutlook key set <api-key> | key test [<api-key>] | key clear");
        }
    }

    // ---------------------------------------------------------------------
    // playbooks / rules
    // ---------------------------------------------------------------------

    private static int Playbooks()
    {
        foreach (var playbook in JevOutlook.Rules.Playbooks.All)
        {
            Console.WriteLine($"{playbook.Id,-28} {playbook.Name}");
            Console.WriteLine($"{string.Empty,-28} {playbook.Summary}");
            Console.WriteLine($"{string.Empty,-28} categories: {string.Join(", ", playbook.Rules.Select(r => r.Spam ? r.Name + "*" : r.Name))}");
            Console.WriteLine();
        }
        Console.WriteLine("* archive eligible. Load one with: jevoutlook rules use <playbook-id>");
        return 0;
    }

    private static int Rules(string[] args)
    {
        var sub = args.Length > 0 ? args[0] : "show";
        switch (sub)
        {
            case "show":
                PrintRules(RuleStore.Load());
                return 0;
            case "path":
                Console.WriteLine(AppPaths.Rules);
                return 0;
            case "reset":
                PrintRules(RuleStore.Reset());
                Console.WriteLine("Rules reset to the Universal inbox playbook.");
                return 0;
            case "use":
            {
                if (args.Length < 2) throw new ArgumentException("Usage: jevoutlook rules use <playbook-id>");
                var playbook = JevOutlook.Rules.Playbooks.Find(args[1]) ?? throw new ArgumentException($"Unknown playbook '{args[1]}'. Run 'jevoutlook playbooks'.");
                PrintRules(RuleStore.Save(playbook.Rules));
                Console.WriteLine($"Rules set to playbook '{playbook.Name}'.");
                return 0;
            }
            case "import":
            {
                if (args.Length < 2) throw new ArgumentException("Usage: jevoutlook rules import <file.json>");
                var rules = JsonSerializer.Deserialize<List<LabelRule>>(File.ReadAllText(args[1]), JsonStore.Options)
                    ?? throw new ArgumentException("The file does not contain a JSON array of rules.");
                PrintRules(RuleStore.Save(rules));
                Console.WriteLine("Rules imported.");
                return 0;
            }
            case "export":
            {
                if (args.Length < 2) throw new ArgumentException("Usage: jevoutlook rules export <file.json>");
                File.WriteAllText(args[1], JsonSerializer.Serialize(RuleStore.Load(), JsonStore.Options));
                Console.WriteLine($"Rules written to {args[1]}.");
                return 0;
            }
            default:
                throw new ArgumentException("Usage: jevoutlook rules show | use <playbook-id> | import <file> | export <file> | reset | path");
        }
    }

    private static void PrintRules(IReadOnlyList<LabelRule> rules)
    {
        foreach (var rule in rules)
        {
            Console.WriteLine($"{rule.Name}{(rule.Spam ? "  [archive eligible]" : string.Empty)}");
            Console.WriteLine($"    {rule.Description}");
        }
        Console.WriteLine();
        Console.WriteLine($"{rules.Count} categor{(rules.Count == 1 ? "y" : "ies")}. Edit {AppPaths.Rules} or use 'rules import' to customise.");
    }

    // ---------------------------------------------------------------------
    // run / continue / status
    // ---------------------------------------------------------------------

    private static async Task<int> RunAsync(string[] args, CancellationToken ct)
    {
        var options = RunOptions.Normalize(
            scope: Option(args, "--scope"),
            unreadOnly: !args.Contains("--include-read"),
            dryRun: !args.Contains("--live"),
            mode: Option(args, "--mode"),
            limitRaw: Option(args, "--limit"),
            maxSpendUsd: Number(args, "--max-spend", AppConstants.DefaultMaxSpendUsd),
            metadataThreshold: Number(args, "--metadata-threshold", AppConstants.DefaultMetadataThreshold),
            archiveThreshold: Number(args, "--archive-threshold", AppConstants.DefaultArchiveThreshold));

        var config = AppConfig.Load();
        var apiKey = config.ResolveApiKey(Option(args, "--api-key"));
        var rules = RuleStore.Load();

        if (!options.DryRun && options.Mode == RunMode.LabelsArchive && !args.Contains("--yes"))
        {
            var eligible = string.Join(", ", rules.Where(r => r.Spam).Select(r => r.Name));
            Console.WriteLine($"LIVE run with archiving: messages classified as [{eligible}] with confidence ≥ {options.ArchiveThreshold.ToString("0.00", CultureInfo.InvariantCulture)} will be moved to the Archive folder.");
            Console.Write("Type 'yes' to continue: ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "yes", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Aborted.");
                return 1;
            }
        }

        var graph = await ConnectGraphAsync(config, ct);
        var sink = new ConsoleSink();
        var engine = new TriageEngine(graph, new JevClient(Http, config.ResolvedEndpoint, apiKey), config.ResolvedModel, sink);

        Console.WriteLine($"{(options.DryRun ? "PREVIEW" : "LIVE")} · scope {options.Scope}{(options.UnreadOnly ? " (unread only)" : string.Empty)} · limit {(options.Limit is { } l ? l.ToString(CultureInfo.InvariantCulture) : "all")} · " +
                          $"{(options.Mode == RunMode.LabelsArchive ? "categories + archive" : "categories only")} · budget ${options.MaxSpendUsd.ToString("0.00", CultureInfo.InvariantCulture)} · model {config.ResolvedModel}");
        Console.WriteLine("Press Ctrl+C to pause safely at any time.");

        var job = await engine.StartAsync(options, rules, ct);
        await engine.RunLoopAsync(job, ct);
        ConsoleSink.PrintSummary(job);
        return job.Status is JobStatus.Completed ? 0 : job.Status == JobStatus.Error ? 1 : 3;
    }

    private static async Task<int> ContinueAsync(string[] args, CancellationToken ct)
    {
        var config = AppConfig.Load();
        var apiKey = config.ResolveApiKey(Option(args, "--api-key"));
        var newSpend = Option(args, "--max-spend") is { } raw
            ? double.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture)
            : (double?)null;

        var job = TriageEngine.Resume(newSpend);
        var graph = await ConnectGraphAsync(config, ct);
        var sink = new ConsoleSink();
        var engine = new TriageEngine(graph, new JevClient(Http, config.ResolvedEndpoint, apiKey), config.ResolvedModel, sink);
        Console.WriteLine($"Continuing session {job.Id} ({(job.DryRun ? "PREVIEW" : "LIVE")}). Press Ctrl+C to pause safely.");
        await engine.RunLoopAsync(job, ct);
        ConsoleSink.PrintSummary(job);
        return job.Status is JobStatus.Completed ? 0 : job.Status == JobStatus.Error ? 1 : 3;
    }

    private static int Status()
    {
        var job = JobStore.Load();
        if (job is null) { Console.WriteLine("No processing session. Start one with: jevoutlook run"); return 0; }
        ConsoleSink.PrintSummary(job);
        if (job.Pending.Count > 0) Console.WriteLine($"  Pending       : {job.Pending.Count} message(s) checkpointed for the next batch");
        return 0;
    }

    private static int ClearJob()
    {
        var job = JobStore.Load();
        if (job is { Status: JobStatus.Running }) throw new InvalidOperationException("Stop processing before clearing the session.");
        JobStore.Delete();
        Console.WriteLine("Session cleared.");
        return 0;
    }

    // ---------------------------------------------------------------------
    // argument helpers
    // ---------------------------------------------------------------------

    private static string? Option(string[] args, string name)
    {
        for (var i = 0; i < args.Length; i++)
        {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
            {
                return i + 1 < args.Length ? args[i + 1] : throw new ArgumentException($"Option {name} requires a value.");
            }
            if (args[i].StartsWith(name + "=", StringComparison.OrdinalIgnoreCase)) return args[i][(name.Length + 1)..];
        }
        return null;
    }

    private static double Number(string[] args, string name, double fallback)
    {
        var raw = Option(args, name);
        if (raw is null) return fallback;
        return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : throw new ArgumentException($"Option {name} expects a number (got '{raw}').");
    }
}
