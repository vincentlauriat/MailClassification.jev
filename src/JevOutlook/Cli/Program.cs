using System.Globalization;
using System.Text;
using System.Text.Json;
using JevOutlook.Graph;
using JevOutlook.Imap;
using JevOutlook.Jev;
using JevOutlook.Mail;
using JevOutlook.Rules;
using JevOutlook.Storage;
using JevOutlook.Triage;

namespace JevOutlook.Cli;

public static class Program
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(100) };

    public static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };
        using var sigterm = System.Runtime.InteropServices.PosixSignalRegistration.Create(
            System.Runtime.InteropServices.PosixSignal.SIGTERM, ctx => { ctx.Cancel = true; cts.Cancel(); });

        try
        {
            var migrated = await GraphTokenProvider.MigrateLegacyAsync(cts.Token);
            if (migrated is not null) Console.Error.WriteLine($"Migrated the existing Microsoft sign-in into account '{migrated.Id}' ({migrated.Email}).");

            var command = args.Length > 0 ? args[0].ToLowerInvariant() : "help";
            var rest = args.Skip(1).ToArray();
            return command switch
            {
                "help" or "--help" or "-h" => Help(),
                "version" or "--version" => Version(),
                "account" or "accounts" => await AccountAsync(rest, cts.Token),
                "auth" => await AuthAsync(rest, cts.Token),
                "config" => Config(rest),
                "key" => await KeyAsync(rest, cts.Token),
                "playbooks" => Playbooks(),
                "rules" => Rules(rest),
                "ui" or "dashboard" or "serve" => await UiAsync(rest, cts.Token),
                "addin" => AddIn(rest),
                "service" => await ServiceCommand.RunAsync(rest, Http, cts.Token),
                "cleanup" => await CleanupAsync(rest, cts.Token),
                "run" => await RunAsync(rest, cts.Token),
                "continue" or "resume" => await ContinueAsync(rest, cts.Token),
                "status" => Status(rest),
                "clear-job" or "clear" => ClearJob(rest),
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
            jevoutlook — AI-assisted mailbox classification with Jev (TypeSafe) via OpenRouter
            Works with Microsoft 365 / Outlook.com (Graph), Gmail (IMAP + app password) and any IMAP server.

            USAGE
              jevoutlook <command> [options]

            MAILBOXES
              account add <email> --m365 [--tenant <id>] [--device-code]
                                                 Microsoft 365 / Outlook.com (signs in right away)
              account add <email> --gmail        Gmail over IMAP (asks for a Google app password)
              account add <email> --imap <host[:port]> [--username <login>]
                                                 Any IMAP server (asks for the mailbox password)
              account list                       Configured mailboxes and their state
              account test <id>                  Connect, show identity, folders and label support
              account login <id> [--device-code] Sign in again (Microsoft 365)
              account password <id>              Replace the stored password (IMAP / Gmail)
              account remove <id>                Forget a mailbox (and its password / sign-in)
              Most commands take --account <id|email>; it is optional with a single mailbox.

            SETUP
              config set client-id <guid>        Entra ID app registration (client) id (Microsoft 365 only)
              config set tenant-id <id>          common (default) | organizations | consumers | <tenant guid>
              config set provider <name>         openrouter (default) | typesafe
              config set device-code true|false  Use the device-code flow instead of a browser
              config show
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

            DASHBOARD
              ui [--port 5177] [--https-port 5178] [--no-https] [--no-open]
                                                 Local web dashboard for all mailboxes
              addin manifest [--out file.xml]    Outlook add-in manifest (Microsoft 365 only)

            SERVICE (macOS)
              service install [--port 5177] [--https-port 5178] [--exe <path>]
                                                 Keep 'ui --no-open' running from login (LaunchAgent)
              service status [--port 5177]       Agent state, /health probe and log path
              service restart                    Restart the agent (after a rebuild)
              service uninstall                  Stop the agent and remove it

            MAINTENANCE
              cleanup remove-category <name> [--account <id>] [--dry-run] [--keep-master]
                                                 Remove a label from every message carrying it, then delete it

            PROCESSING
              run [options] [--account <id>]     Start a session (PREVIEW by default)
                --all-accounts                   Run the same session on every configured mailbox, one after another
                --scope inbox|all                Inbox (default) or the whole mailbox (Outlook, Gmail)
                --include-read                   Also process read messages (default: unread only)
                --limit N|all                    Number of messages (default 10)
                --mode labels|labels-archive     Labels only (default) or archive eligible labels
                --live                           Apply changes to the mailbox (default: preview only)
                --max-spend <usd>                Run budget per mailbox (default 0.10)
                --metadata-threshold <0.5-0.99>  Confidence required to trust the metadata pass (default 0.75)
                --archive-threshold <0.5-0.999>  Confidence required to archive (default 0.93)
                --yes                            Skip the confirmation for live archive runs
              continue [--max-spend <usd>] [--account <id>]
                                                 Resume a paused session (raise the budget if needed)
              status [--account <id>]            Show the current session(s)
              clear-job [--account <id>]         Forget a finished session

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
    // accounts
    // ---------------------------------------------------------------------

    private static async Task<int> AccountAsync(string[] args, CancellationToken ct)
    {
        var sub = args.Length > 0 ? args[0].ToLowerInvariant() : "list";
        switch (sub)
        {
            case "list":
            {
                var accounts = AccountStore.Load();
                if (accounts.Count == 0) { Console.WriteLine("No mailbox configured. Add one with: jevoutlook account add <email> --m365 | --gmail | --imap <host>"); return 0; }
                foreach (var a in accounts)
                {
                    var job = new JobStore(a.Id).Load();
                    var state = MailboxFactory.IsReady(a) ? "ready" : a.IsGraph ? "not signed in" : "no password";
                    var where = a.IsImap ? $" · {a.Host}:{a.Port}" : string.Empty;
                    Console.WriteLine($"{a.Id,-32} {a.Email,-36} {a.ProviderLabel,-13} {state}{where}{(job is null ? string.Empty : $" · session {job.Status}")}");
                }
                Console.WriteLine();
                Console.WriteLine($"Passwords are kept in the {SecretStore.Backend}. State directory: {AppPaths.Root}");
                return 0;
            }
            case "add":
                return await AccountAddAsync(args.Skip(1).ToArray(), ct);
            case "login":
            {
                var account = AccountStore.Resolve(args.Length > 1 && !args[1].StartsWith("--") ? args[1] : null);
                if (!account.IsGraph) throw new ArgumentException($"'{account.Id}' is an IMAP mailbox; use 'account password {account.Id}' to replace its password.");
                await SignInGraphAsync(account, args.Contains("--device-code"), ct);
                return 0;
            }
            case "password":
            {
                var account = AccountStore.Resolve(args.Length > 1 ? args[1] : null);
                if (!account.IsImap) throw new ArgumentException($"'{account.Id}' is a Microsoft 365 mailbox; use 'account login {account.Id}'.");
                var password = ReadPassword(account.Gmail ? "Gmail app password: " : "Mailbox password: ");
                await VerifyImapAsync(account, password, ct);
                SecretStore.Set(account.Id, password);
                Console.WriteLine($"Password stored in the {SecretStore.Backend}.");
                return 0;
            }
            case "test":
            {
                var account = AccountStore.Resolve(args.Length > 1 ? args[1] : null);
                var mailbox = await MailboxFactory.OpenAsync(account, AppConfig.Load(), Http, ct);
                Console.WriteLine($"Identity      : {await mailbox.GetIdentityAsync(ct)}");
                var caps = await mailbox.ConnectAsync(ct);
                Console.WriteLine($"Provider      : {caps.ProviderName} ({caps.LabelNounPlural}; archive {(caps.CanArchive ? "available" : "unavailable")}; whole-mailbox scope {(caps.SupportsAllScope ? "supported" : "not supported")})");
                Console.WriteLine($"Archive       : {caps.ArchiveDescription}");
                Console.WriteLine($"Inbox         : {await mailbox.EstimateAsync("inbox", false, ct)} messages, {await mailbox.EstimateAsync("inbox", true, ct)} unread");
                return 0;
            }
            case "remove":
            {
                if (args.Length < 2) throw new ArgumentException("Usage: jevoutlook account remove <id>");
                var account = AccountStore.Require(args[1]);
                if (new JobStore(account.Id).Load() is { Status: JobStatus.Running }) throw new InvalidOperationException("Stop processing on this mailbox before removing it.");
                AccountStore.Remove(account.Id);
                Console.WriteLine($"Mailbox '{account.Id}' removed (state directory, sign-in record and password deleted).");
                return 0;
            }
            default:
                throw new ArgumentException("Usage: jevoutlook account list | add <email> (--m365 | --gmail | --imap <host[:port]>) | login <id> | password <id> | test <id> | remove <id>");
        }
    }

    private static async Task<int> AccountAddAsync(string[] args, CancellationToken ct)
    {
        var email = args.FirstOrDefault(a => !a.StartsWith("--") && a.Contains('@'))
            ?? throw new ArgumentException("Usage: jevoutlook account add <email> --m365 | --gmail | --imap <host[:port]>");
        var account = new MailAccount { Id = MailAccount.IdFor(email), Email = email.Trim() };

        if (args.Contains("--m365") || args.Contains("--graph") || args.Contains("--outlook"))
        {
            account.Kind = AccountKind.Graph;
            if (Option(args, "--tenant") is { Length: > 0 } tenant) account.TenantId = tenant;
            AccountStore.Add(account);
            try
            {
                await SignInGraphAsync(account, args.Contains("--device-code"), ct);
            }
            catch
            {
                AccountStore.Remove(account.Id);
                throw;
            }
            return 0;
        }

        if (args.Contains("--gmail"))
        {
            account.Kind = AccountKind.Imap;
            account.Gmail = true;
            account.Host = "imap.gmail.com";
            account.Port = 993;
            Console.WriteLine("Gmail over IMAP needs a Google app password (2-step verification required): https://myaccount.google.com/apppasswords");
        }
        else if (Option(args, "--imap") is { Length: > 0 } host)
        {
            account.Kind = AccountKind.Imap;
            var parts = host.Split(':', 2);
            account.Host = parts[0].Trim();
            account.Port = parts.Length > 1 && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var port) ? port : 993;
            if (Option(args, "--username") is { Length: > 0 } user && !string.Equals(user, email, StringComparison.OrdinalIgnoreCase)) account.Username = user;
        }
        else
        {
            throw new ArgumentException("Choose the mailbox type: --m365, --gmail or --imap <host[:port]>.");
        }

        var password = Option(args, "--password-stdin") is not null || args.Contains("--password-stdin")
            ? (Console.In.ReadLine() ?? string.Empty).Trim()
            : ReadPassword(account.Gmail ? "Gmail app password: " : "Mailbox password: ");
        await VerifyImapAsync(account, password, ct);
        AccountStore.Add(account);
        SecretStore.Set(account.Id, password);
        Console.WriteLine($"Mailbox '{account.Id}' added ({account.ProviderLabel}, {account.Host}:{account.Port}). Password stored in the {SecretStore.Backend}.");
        return 0;
    }

    private static async Task VerifyImapAsync(MailAccount account, string password, CancellationToken ct)
    {
        if (password.Length == 0) throw new ArgumentException("The password is empty.");
        var probe = new ImapMailbox(account, password);
        Console.WriteLine($"Connecting to {account.Host}:{account.Port}…");
        var identity = await probe.GetIdentityAsync(ct);
        var caps = await probe.ConnectAsync(ct);
        Console.WriteLine($"Connected as {identity} ({caps.ProviderName}; {caps.LabelNounPlural}; archive {(caps.CanArchive ? "available" : "unavailable")}).");
    }

    private static async Task SignInGraphAsync(MailAccount account, bool deviceCode, CancellationToken ct)
    {
        var config = AppConfig.Load();
        Console.WriteLine($"Opening the Microsoft sign-in flow for {account.Email}…");
        var mailbox = await MailboxFactory.SignInGraphAsync(account, config, Http, deviceCode, null, ct);
        var accounts = AccountStore.Load();
        var saved = accounts.First(a => a.Id == account.Id);
        saved.Email = account.Email;
        AccountStore.Save(accounts);
        Console.WriteLine("Signed in as " + await mailbox.GetIdentityAsync(ct));
    }

    /// <summary>Hidden password prompt (echo off); falls back to a plain line when no terminal is attached.</summary>
    private static string ReadPassword(string prompt)
    {
        Console.Write(prompt);
        if (Console.IsInputRedirected) return (Console.ReadLine() ?? string.Empty).Trim();
        var sb = new StringBuilder();
        while (true)
        {
            var key = Console.ReadKey(intercept: true);
            if (key.Key == ConsoleKey.Enter) { Console.WriteLine(); break; }
            if (key.Key == ConsoleKey.Backspace) { if (sb.Length > 0) sb.Length--; continue; }
            if (!char.IsControl(key.KeyChar)) sb.Append(key.KeyChar);
        }
        return sb.ToString().Trim();
    }

    private static MailAccount ResolveAccount(string[] args) => AccountStore.Resolve(Option(args, "--account"));

    // ---------------------------------------------------------------------
    // config / auth (compatibility)
    // ---------------------------------------------------------------------

    private static int Config(string[] args)
    {
        var config = AppConfig.Load();
        if (args.Length == 0 || args[0] == "show")
        {
            Console.WriteLine($"State directory : {AppPaths.Root}");
            Console.WriteLine($"Mailboxes       : {AccountStore.Load().Count} (jevoutlook account list)");
            Console.WriteLine($"Client id       : {config.ClientId ?? "(not set)"}");
            Console.WriteLine($"Tenant id       : {config.TenantId}");
            Console.WriteLine($"Provider        : {config.Provider} → {config.ResolvedEndpoint}");
            Console.WriteLine($"Model           : {config.ResolvedModel}");
            Console.WriteLine($"Device code     : {config.DeviceCode}");
            Console.WriteLine($"Stored API key  : {(string.IsNullOrEmpty(config.ApiKey) ? "no" : "yes")}");
            Console.WriteLine($"Secrets         : {SecretStore.Backend}");
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

    /// <summary>Kept for muscle memory: 'auth' is 'account login' on the selected (or only) mailbox.</summary>
    private static async Task<int> AuthAsync(string[] args, CancellationToken ct)
    {
        var account = ResolveAccount(args);
        if (args.Contains("--logout"))
        {
            if (account.IsGraph) GraphTokenProvider.DeleteRecord(account.Id); else SecretStore.Delete(account.Id);
            Console.WriteLine($"Signed out of '{account.Id}'.");
            return 0;
        }
        if (args.Contains("--status"))
        {
            if (!MailboxFactory.IsReady(account)) { Console.WriteLine($"'{account.Id}' is not signed in. Run: jevoutlook account login {account.Id}"); return 1; }
            var mailbox = await MailboxFactory.OpenAsync(account, AppConfig.Load(), Http, ct);
            Console.WriteLine("Signed in as " + await mailbox.GetIdentityAsync(ct));
            return 0;
        }
        if (!account.IsGraph) throw new ArgumentException($"'{account.Id}' is an IMAP mailbox; use 'account password {account.Id}'.");
        await SignInGraphAsync(account, args.Contains("--device-code"), ct);
        return 0;
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
    // web dashboard
    // ---------------------------------------------------------------------

    private static async Task<int> UiAsync(string[] args, CancellationToken ct)
    {
        var port = (int)Number(args, "--port", 5177);
        var httpsPort = (int)Number(args, "--https-port", 5178);
        if (port is < 1 or > 65535 || httpsPort is < 1 or > 65535) throw new ArgumentException("Ports must be between 1 and 65535.");

        // Single instance: a second 'ui' (by hand while the LaunchAgent runs, or the
        // agent while a manual one runs) reuses the running server instead of failing to bind.
        var (state, version) = await Web.UiServer.ProbeAsync(Http, port, ct);
        if (state == Web.UiServer.HealthState.JevOutlook)
        {
            var url = $"http://127.0.0.1:{port}/";
            Console.WriteLine($"jevOutlook {version} is already running: {url}");
            if (!args.Contains("--no-open")) Web.UiServer.TryOpenBrowser(url);
            return 0;
        }
        if (state == Web.UiServer.HealthState.Foreign)
        {
            throw new InvalidOperationException($"Port {port} answers but not as a current jevOutlook (another program, or an older jevOutlook without /health). Stop it, or pick another port with --port.");
        }

        await new Web.UiServer(Http).RunAsync(port, args.Contains("--no-https") ? null : httpsPort, openBrowser: !args.Contains("--no-open"), ct);
        return 0;
    }

    /// <summary>Write the Outlook add-in manifest to release/ for sideloading.</summary>
    private static int AddIn(string[] args)
    {
        if (args.Length == 0 || args[0] != "manifest") throw new ArgumentException("Usage: jevoutlook addin manifest [--https-port 5178] [--out <file.xml>]");
        var httpsPort = (int)Number(args, "--https-port", 5178);
        var output = Option(args, "--out") ?? Path.Combine(Directory.GetCurrentDirectory(), "release", "jevoutlook-manifest.xml");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        File.WriteAllText(output, Web.UiServer.BuildManifest(httpsPort));
        Console.WriteLine($"Manifest written to {output}");
        Console.WriteLine("Sideload it in Outlook: Get Add-ins → My add-ins → Add a custom add-in → Add from file.");
        Console.WriteLine("Run 'jevoutlook service install' once so the pane is always available (macOS), or keep 'jevoutlook ui' running.");
        return 0;
    }

    // ---------------------------------------------------------------------
    // maintenance
    // ---------------------------------------------------------------------

    private static async Task<int> CleanupAsync(string[] args, CancellationToken ct)
    {
        if (args.Length < 2 || args[0] != "remove-category" || string.IsNullOrWhiteSpace(args[1]))
            throw new ArgumentException("Usage: jevoutlook cleanup remove-category <name> [--account <id>] [--dry-run] [--keep-master]");
        var category = args[1].Trim();
        var dryRun = args.Contains("--dry-run");
        await using var mailbox = await MailboxFactory.OpenAsync(ResolveAccount(args), AppConfig.Load(), Http, ct);
        await mailbox.ConnectAsync(ct);

        var total = 0;
        for (var rounds = 0; rounds < 500; rounds++)
        {
            ct.ThrowIfCancellationRequested();
            var found = await mailbox.FindMessagesWithLabelAsync(category, 100, ct);
            if (found.Count == 0) break;
            if (dryRun)
            {
                Console.WriteLine($"{found.Count}{(found.Count == 100 ? "+" : string.Empty)} message(s) carry “{category}”. Nothing changed (dry run).");
                return 0;
            }
            var outcome = await mailbox.RemoveLabelAsync(found, category, ct);
            if (!outcome.Ok) throw new InvalidOperationException(outcome.Error);
            total += found.Count;
            Console.WriteLine($"Removed “{category}” from {total} message(s)…");
            if (found.Count < 100) break;
        }
        Console.WriteLine(total == 0 ? $"No message carries “{category}”." : $"Done: “{category}” removed from {total} message(s).");

        if (!dryRun && !args.Contains("--keep-master"))
        {
            Console.WriteLine(await mailbox.DeleteLabelAsync(category, ct)
                ? $"“{category}” deleted from the label list."
                : $"“{category}” was not in the label list.");
        }
        return 0;
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
        var accounts = args.Contains("--all-accounts") ? AccountStore.Load() : [ResolveAccount(args)];
        if (accounts.Count == 0) throw new InvalidOperationException("No mailbox is configured. Add one with: jevoutlook account add <email> …");

        if (!options.DryRun && options.Mode == RunMode.LabelsArchive && !args.Contains("--yes"))
        {
            var eligible = string.Join(", ", rules.Where(r => r.Spam).Select(r => r.Name));
            Console.WriteLine($"LIVE run with archiving on {string.Join(", ", accounts.Select(a => a.Email))}: messages classified as [{eligible}] with confidence ≥ {options.ArchiveThreshold.ToString("0.00", CultureInfo.InvariantCulture)} will be archived.");
            Console.Write("Type 'yes' to continue: ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "yes", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Aborted.");
                return 1;
            }
        }

        var worst = 0;
        foreach (var account in accounts)
        {
            ct.ThrowIfCancellationRequested();
            if (accounts.Count > 1) Console.WriteLine($"═══ {account.Email} ({account.ProviderLabel}) ═══");
            // Held for the whole run loop, released on exit, exception or Ctrl+C.
            var store = new JobStore(account.Id);
            using var jobLock = store.TryLock();
            if (jobLock is null)
            {
                if (accounts.Count == 1) throw new InvalidOperationException(JobStore.LockedMessage);
                Console.Error.WriteLine($"Skipped {account.Email}: {JobStore.LockedMessage}");
                worst = Math.Max(worst, 1);
                continue;
            }
            await using var mailbox = await MailboxFactory.OpenAsync(account, config, Http, ct);
            var sink = new ConsoleSink();
            var engine = new TriageEngine(mailbox, store, new JevClient(Http, config.ResolvedEndpoint, apiKey), config.ResolvedModel, sink);

            Console.WriteLine($"{(options.DryRun ? "PREVIEW" : "LIVE")} · {account.Email} · scope {options.Scope}{(options.UnreadOnly ? " (unread only)" : string.Empty)} · limit {(options.Limit is { } l ? l.ToString(CultureInfo.InvariantCulture) : "all")} · " +
                              $"{(options.Mode == RunMode.LabelsArchive ? "labels + archive" : "labels only")} · budget ${options.MaxSpendUsd.ToString("0.00", CultureInfo.InvariantCulture)} · model {config.ResolvedModel}");
            Console.WriteLine("Press Ctrl+C to pause safely at any time.");

            var job = await engine.StartAsync(options, rules, ct);
            await engine.RunLoopAsync(job, ct);
            ConsoleSink.PrintSummary(job);
            var code = job.Status is JobStatus.Completed ? 0 : job.Status == JobStatus.Error ? 1 : 3;
            worst = Math.Max(worst, code);
        }
        return worst;
    }

    private static async Task<int> ContinueAsync(string[] args, CancellationToken ct)
    {
        var config = AppConfig.Load();
        var apiKey = config.ResolveApiKey(Option(args, "--api-key"));
        var newSpend = Option(args, "--max-spend") is { } raw
            ? double.Parse(raw, NumberStyles.Float, CultureInfo.InvariantCulture)
            : (double?)null;

        var account = ResolveAccount(args);
        // Taken before Resume (which rewrites job.json) and held for the whole run loop.
        var store = new JobStore(account.Id);
        using var jobLock = store.Lock();
        await using var mailbox = await MailboxFactory.OpenAsync(account, config, Http, ct);
        var sink = new ConsoleSink();
        var engine = new TriageEngine(mailbox, store, new JevClient(Http, config.ResolvedEndpoint, apiKey), config.ResolvedModel, sink);
        var job = engine.Resume(newSpend);
        Console.WriteLine($"Continuing session {job.Id} on {account.Email} ({(job.DryRun ? "PREVIEW" : "LIVE")}). Press Ctrl+C to pause safely.");
        await engine.RunLoopAsync(job, ct);
        ConsoleSink.PrintSummary(job);
        return job.Status is JobStatus.Completed ? 0 : job.Status == JobStatus.Error ? 1 : 3;
    }

    private static int Status(string[] args)
    {
        var accounts = Option(args, "--account") is { } id ? [AccountStore.Require(id)] : AccountStore.Load();
        if (accounts.Count == 0) { Console.WriteLine("No mailbox configured."); return 0; }
        var any = false;
        foreach (var account in accounts)
        {
            var job = new JobStore(account.Id).Load();
            if (job is null) continue;
            any = true;
            Console.WriteLine($"═══ {account.Email} ({account.ProviderLabel}) ═══");
            ConsoleSink.PrintSummary(job);
            if (job.Pending.Count > 0) Console.WriteLine($"  Pending       : {job.Pending.Count} message(s) checkpointed for the next batch");
        }
        if (!any) Console.WriteLine("No processing session. Start one with: jevoutlook run");
        return 0;
    }

    private static int ClearJob(string[] args)
    {
        var account = ResolveAccount(args);
        var store = new JobStore(account.Id);
        using var jobLock = store.Lock();
        var job = store.Load();
        if (job is { Status: JobStatus.Running }) throw new InvalidOperationException("Stop processing before clearing the session.");
        store.Delete();
        Console.WriteLine($"Session cleared for {account.Email}.");
        return 0;
    }

    // ---------------------------------------------------------------------
    // argument helpers
    // ---------------------------------------------------------------------

    internal static string? Option(string[] args, string name)
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

    internal static double Number(string[] args, string name, double fallback)
    {
        var raw = Option(args, name);
        if (raw is null) return fallback;
        return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : throw new ArgumentException($"Option {name} expects a number (got '{raw}').");
    }
}
