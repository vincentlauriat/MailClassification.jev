using System.Globalization;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using Azure.Identity;
using JevOutlook.Graph;
using JevOutlook.Jev;
using JevOutlook.Mail;
using JevOutlook.Rules;
using JevOutlook.Storage;
using JevOutlook.Triage;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace JevOutlook.Web;

/// <summary>
/// Local web dashboard: the counterpart of jevMail's Apps Script web app, for
/// several mailboxes (Microsoft 365 via Graph, Gmail and generic IMAP). The page
/// (embedded <c>wwwroot/index.html</c>) keeps the original <c>google.script.run</c>
/// calling convention; a small bridge turns every call into <c>POST /api/{function}</c>
/// with the arguments as a JSON array. Listens on 127.0.0.1 only.
/// </summary>
public sealed class UiServer
{
    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>Stable identity of the Outlook add-in (manifest &lt;Id&gt;).</summary>
    public const string AddInId = "7c1f3f0e-6d2a-4b5e-9c1a-2f0e8a5d4b31";
    public const string AddInVersion = "1.0.0.0";

    private readonly HttpClient _http;
    private readonly SemaphoreSlim _jobLock = new(1, 1);
    private int _httpPort;
    private int? _httpsPort;

    // Microsoft sign-in state shared with the page (device-code flow), per account.
    private sealed class SignInState
    {
        public Task? Task;
        public string? Code, Url, Error;
    }
    private readonly object _authGate = new();
    private readonly Dictionary<string, SignInState> _signIns = new(StringComparer.OrdinalIgnoreCase);

    public UiServer(HttpClient http) => _http = http;

    public async Task RunAsync(int port, int? httpsPort, bool openBrowser, CancellationToken ct)
    {
        _httpPort = port;
        _httpsPort = httpsPort;
        var app = Build(port, httpsPort);
        try
        {
            await app.StartAsync(ct);
        }
        catch (Exception ex) when (httpsPort is not null && ex is not OperationCanceledException && IsCertificateProblem(ex))
        {
            Console.Error.WriteLine("HTTPS could not be started (no ASP.NET Core development certificate). The Outlook add-in needs it:");
            Console.Error.WriteLine("  dotnet dev-certs https --trust");
            Console.Error.WriteLine("Continuing with HTTP only (dashboard works, add-in pane will not load).");
            _httpsPort = null;
            app = Build(port, null);
            await app.StartAsync(ct);
        }

        var url = $"http://127.0.0.1:{port}/";
        Console.WriteLine($"jevOutlook dashboard: {url}   (Ctrl+C to stop)");
        if (_httpsPort is { } hp)
        {
            Console.WriteLine($"Outlook add-in pane: https://localhost:{hp}/taskpane.html · manifest: https://localhost:{hp}/manifest.xml");
        }
        if (openBrowser) TryOpenBrowser(url);
        try { await Task.Delay(Timeout.Infinite, ct); } catch (OperationCanceledException) { }
        await app.StopAsync(CancellationToken.None);
    }

    private static bool IsCertificateProblem(Exception ex)
    {
        for (var e = ex; e is not null; e = e.InnerException)
        {
            if (e.Message.Contains("certificate", StringComparison.OrdinalIgnoreCase) || e.Message.Contains("HTTPS", StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private WebApplication Build(int port, int? httpsPort)
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions { Args = [] });
        builder.Logging.ClearProviders();
        builder.Logging.AddSimpleConsole();
        builder.Logging.SetMinimumLevel(LogLevel.Warning);
        builder.WebHost.ConfigureKestrel(kestrel =>
        {
            kestrel.ListenLocalhost(port);
            if (httpsPort is { } hp) kestrel.ListenLocalhost(hp, listen => listen.UseHttps());
        });
        var app = builder.Build();

        // Defense against DNS rebinding and cross-site requests to the local API:
        // only our own Host, only same-origin API calls, only JSON bodies (forces a
        // CORS preflight, which this server never answers).
        var (allowedHosts, allowedOrigins) = AllowedOrigins(port, httpsPort);
        app.Use(async (context, next) =>
        {
            if (!allowedHosts.Contains(context.Request.Host.Value ?? string.Empty))
            {
                context.Response.StatusCode = StatusCodes.Status421MisdirectedRequest;
                return;
            }
            if (context.Request.Path.StartsWithSegments("/api"))
            {
                var origin = context.Request.Headers.Origin.ToString();
                var fetchSite = context.Request.Headers["Sec-Fetch-Site"].ToString();
                var sameOrigin = (origin.Length == 0 || allowedOrigins.Contains(origin)) &&
                                 (fetchSite.Length == 0 || fetchSite is "same-origin" or "none");
                var isJson = context.Request.ContentType?.StartsWith("application/json", StringComparison.OrdinalIgnoreCase) == true;
                if (!sameOrigin || !isJson || !HttpMethods.IsPost(context.Request.Method))
                {
                    context.Response.StatusCode = StatusCodes.Status403Forbidden;
                    await context.Response.WriteAsJsonAsync(new { error = "Cross-site or non-JSON requests to the local API are rejected." });
                    return;
                }
            }
            await next();
        });

        app.MapGet("/", () => Static("index.html", "text/html; charset=utf-8"));
        app.MapGet("/health", () => Results.Json(new { app = HealthAppName, version = AppVersion, https = httpsPort is not null }, Json));
        app.MapGet("/manifest.xml", () => httpsPort is { } hp
            ? Results.Text(BuildManifest(hp), "application/xml; charset=utf-8")
            : Results.Text("HTTPS is not enabled; start the server with an HTTPS port to get the add-in manifest.", "text/plain", statusCode: 503));
        app.MapGet("/{file}", (string file) => Static(file, ContentType(file)));
        app.MapPost("/api/{function}", HandleApiAsync);
        return app;
    }

    /// <summary>
    /// Host headers and Origins the local server accepts (plain port on 127.0.0.1 and
    /// localhost, plus the HTTPS port when enabled). Every allowlist lives here.
    /// </summary>
    public static (HashSet<string> Hosts, HashSet<string> Origins) AllowedOrigins(int port, int? httpsPort)
    {
        var hosts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            $"127.0.0.1:{port}", $"localhost:{port}",
        };
        var origins = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            $"http://127.0.0.1:{port}", $"http://localhost:{port}",
        };
        if (httpsPort is { } hps)
        {
            hosts.Add($"localhost:{hps}");
            hosts.Add($"127.0.0.1:{hps}");
            origins.Add($"https://localhost:{hps}");
            origins.Add($"https://127.0.0.1:{hps}");
        }
        return (hosts, origins);
    }

    // ---------------------------------------------------------------------
    // health (single instance, service status)
    // ---------------------------------------------------------------------

    /// <summary>Value of <c>app</c> in the <c>GET /health</c> answer.</summary>
    public const string HealthAppName = "jevoutlook";

    public static string AppVersion => typeof(UiServer).Assembly.GetName().Version?.ToString(3) ?? "0.0.0";

    public enum HealthState { Free, JevOutlook, Foreign }

    /// <summary>
    /// Probe <c>http://127.0.0.1:{port}/health</c>. <see cref="HealthState.Free"/>: nothing
    /// listens; <see cref="HealthState.JevOutlook"/>: a jevOutlook server answers;
    /// <see cref="HealthState.Foreign"/>: something else (or an older jevOutlook without
    /// <c>/health</c>) holds the port.
    /// </summary>
    public static async Task<(HealthState State, string? Version)> ProbeAsync(HttpClient http, int port, CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(2));
        try
        {
            using var response = await http.GetAsync($"http://127.0.0.1:{port}/health", timeout.Token);
            var body = await response.Content.ReadAsStringAsync(timeout.Token);
            return response.IsSuccessStatusCode && ParseHealth(body) is { } version
                ? (HealthState.JevOutlook, version)
                : (HealthState.Foreign, null);
        }
        catch (HttpRequestException ex) when (ex.InnerException is System.Net.Sockets.SocketException { SocketErrorCode: System.Net.Sockets.SocketError.ConnectionRefused })
        {
            return (HealthState.Free, null); // nothing listens
        }
        catch (HttpRequestException)
        {
            return (HealthState.Foreign, null); // a listener that does not speak HTTP (reset, garbage)
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return (HealthState.Foreign, null); // accepted the connection but never answered
        }
    }

    /// <summary>The version from a jevOutlook <c>/health</c> body, or null when the body is anything else.</summary>
    internal static string? ParseHealth(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("app", out var app) || app.ValueKind != JsonValueKind.String || app.GetString() != HealthAppName) return null;
            return root.TryGetProperty("version", out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : string.Empty;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    // ---------------------------------------------------------------------
    // Outlook add-in manifest
    // ---------------------------------------------------------------------

    /// <summary>Classic XML manifest for a read-mode task pane (sideloadable in Outlook on the web, new Outlook, Mac and Windows).</summary>
    public static string BuildManifest(int httpsPort)
    {
        var b = $"https://localhost:{httpsPort}";
        return $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
                       xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                       xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
                       xsi:type="MailApp">
              <Id>{AddInId}</Id>
              <Version>{AddInVersion}</Version>
              <ProviderName>jevOutlook</ProviderName>
              <DefaultLocale>en-US</DefaultLocale>
              <DisplayName DefaultValue="jevOutlook"/>
              <Description DefaultValue="AI-assisted classification of Outlook messages with Jev (TypeSafe). Runs against the local jevOutlook server."/>
              <IconUrl DefaultValue="{b}/icon-64.png"/>
              <HighResolutionIconUrl DefaultValue="{b}/icon-128.png"/>
              <SupportUrl DefaultValue="https://github.com/vincentlauriat/MailClassification.jev"/>
              <AppDomains>
                <AppDomain>{b}</AppDomain>
              </AppDomains>
              <Hosts>
                <Host Name="Mailbox"/>
              </Hosts>
              <Requirements>
                <Sets>
                  <Set Name="Mailbox" MinVersion="1.1"/>
                </Sets>
              </Requirements>
              <FormSettings>
                <Form xsi:type="ItemRead">
                  <DesktopSettings>
                    <SourceLocation DefaultValue="{b}/taskpane.html"/>
                    <RequestedHeight>450</RequestedHeight>
                  </DesktopSettings>
                </Form>
              </FormSettings>
              <Permissions>ReadWriteItem</Permissions>
              <Rule xsi:type="RuleCollection" Mode="Or">
                <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
              </Rule>
              <DisableEntityHighlighting>false</DisableEntityHighlighting>
              <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
                <Requirements>
                  <bt:Sets DefaultMinVersion="1.3">
                    <bt:Set Name="Mailbox"/>
                  </bt:Sets>
                </Requirements>
                <Hosts>
                  <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                      <FunctionFile resid="Taskpane.Url"/>
                      <ExtensionPoint xsi:type="MessageReadCommandSurface">
                        <OfficeTab id="TabDefault">
                          <Group id="jevOutlookGroup">
                            <Label resid="Group.Label"/>
                            <Control xsi:type="Button" id="jevOutlookOpenPane">
                              <Label resid="Button.Label"/>
                              <Supertip>
                                <Title resid="Button.Label"/>
                                <Description resid="Button.Tooltip"/>
                              </Supertip>
                              <Icon>
                                <bt:Image size="16" resid="Icon.16"/>
                                <bt:Image size="32" resid="Icon.32"/>
                                <bt:Image size="80" resid="Icon.80"/>
                              </Icon>
                              <Action xsi:type="ShowTaskpane">
                                <SourceLocation resid="Taskpane.Url"/>
                              </Action>
                            </Control>
                          </Group>
                        </OfficeTab>
                      </ExtensionPoint>
                    </DesktopFormFactor>
                  </Host>
                </Hosts>
                <Resources>
                  <bt:Images>
                    <bt:Image id="Icon.16" DefaultValue="{b}/icon-16.png"/>
                    <bt:Image id="Icon.32" DefaultValue="{b}/icon-32.png"/>
                    <bt:Image id="Icon.80" DefaultValue="{b}/icon-80.png"/>
                  </bt:Images>
                  <bt:Urls>
                    <bt:Url id="Taskpane.Url" DefaultValue="{b}/taskpane.html"/>
                  </bt:Urls>
                  <bt:ShortStrings>
                    <bt:String id="Group.Label" DefaultValue="jevOutlook"/>
                    <bt:String id="Button.Label" DefaultValue="Classify"/>
                  </bt:ShortStrings>
                  <bt:LongStrings>
                    <bt:String id="Button.Tooltip" DefaultValue="Classify this message with Jev and apply an Outlook category."/>
                  </bt:LongStrings>
                </Resources>
              </VersionOverrides>
            </OfficeApp>
            """;
    }

    // ---------------------------------------------------------------------
    // static assets (embedded)
    // ---------------------------------------------------------------------

    private static IResult Static(string file, string contentType)
    {
        var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("wwwroot/" + file);
        return stream is null ? Results.NotFound() : Results.Stream(stream, contentType);
    }

    private static string ContentType(string file) => Path.GetExtension(file).ToLowerInvariant() switch
    {
        ".html" => "text/html; charset=utf-8",
        ".css" => "text/css; charset=utf-8",
        ".js" => "text/javascript; charset=utf-8",
        ".svg" => "image/svg+xml",
        ".png" => "image/png",
        _ => "application/octet-stream",
    };

    internal static void TryOpenBrowser(string url)
    {
        try
        {
            if (OperatingSystem.IsMacOS()) System.Diagnostics.Process.Start("open", url);
            else if (OperatingSystem.IsWindows()) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
            else System.Diagnostics.Process.Start("xdg-open", url);
        }
        catch (Exception)
        {
            // The URL is printed; the user can open it manually.
        }
    }

    // ---------------------------------------------------------------------
    // API dispatch
    // ---------------------------------------------------------------------

    private async Task<IResult> HandleApiAsync(string function, HttpContext context)
    {
        JsonElement[] args;
        try
        {
            using var doc = await JsonDocument.ParseAsync(context.Request.Body, cancellationToken: context.RequestAborted);
            args = doc.RootElement.ValueKind == JsonValueKind.Array ? doc.RootElement.EnumerateArray().Select(a => a.Clone()).ToArray() : [];
        }
        catch (JsonException)
        {
            args = [];
        }

        try
        {
            object? result = function switch
            {
                "getUiState" => GetUiState(),
                "authStatus" => await AuthStatusAsync(Str(args, 0), context.RequestAborted),
                "startSignIn" => await StartSignInAsync(Str(args, 0)),
                "signOut" => SignOut(Str(args, 0)),
                "addAccount" => await WithJobLockAsync(() => AddAccountAsync(args.Length > 0 ? args[0] : default, context.RequestAborted)),
                "removeAccount" => await WithJobLockAsync(() => Task.FromResult<object?>(RemoveAccount(Str(args, 0)))),
                "testAccount" => await TestAccountAsync(Str(args, 0), context.RequestAborted),
                "testOpenRouterKey" => await TestKeyAsync(Str(args, 0), Bool(args, 1), context.RequestAborted),
                "clearSavedOpenRouterKey" => ClearKey(),
                "saveLabelRules" => await WithJobLockAsync(() => Task.FromResult<object?>(new { ok = true, rules = RuleStore.Save(Rules(args, 0)) })),
                "resetLabelRules" => await WithJobLockAsync(() => Task.FromResult<object?>(new { ok = true, rules = RuleStore.Reset() })),
                "startTriageJob" => await WithJobLockAsync(() => StartJobAsync(args.Length > 0 ? args[0] : default, context.RequestAborted)),
                "processNextBatch" => await WithJobLockAsync(() => ProcessNextBatchAsync(Str(args, 0), Str(args, 1), Str(args, 2), context.RequestAborted)),
                "resumeTriageJob" => await WithJobLockAsync(() => ResumeJobAsync(Str(args, 0), context.RequestAborted)),
                "cancelTriageJob" => await WithJobLockAsync(() => Task.FromResult<object?>(CancelJob(Str(args, 0)))),
                "clearFinishedJob" => await WithJobLockAsync(() => Task.FromResult<object?>(ClearJob(Str(args, 0)))),
                "classifyItem" => await WithJobLockAsync(() => ClassifyItemAsync(Str(args, 0), Str(args, 1), Str(args, 2), context.RequestAborted)),
                "applyItem" => await WithJobLockAsync(() => ApplyItemAsync(Str(args, 0), Str(args, 1), Bool(args, 2), Str(args, 3), context.RequestAborted)),
                _ => throw new InvalidOperationException($"Unknown server function '{function}'."),
            };
            return Results.Json(result, Json);
        }
        catch (OperationCanceledException) when (context.RequestAborted.IsCancellationRequested)
        {
            return Results.StatusCode(499);
        }
        catch (Exception ex)
        {
            return Results.Json(new { error = ex.Message }, Json, statusCode: 500);
        }
    }

    private async Task<object?> WithJobLockAsync(Func<Task<object?>> operation)
    {
        if (!await _jobLock.WaitAsync(TimeSpan.FromSeconds(5)))
            throw new InvalidOperationException("Another batch is still running. Wait for it to finish and try again.");
        try { return await operation(); }
        finally { _jobLock.Release(); }
    }

    // ---------------------------------------------------------------------
    // state / key / rules
    // ---------------------------------------------------------------------

    private object GetUiState()
    {
        var config = AppConfig.Load();
        var hasKey = !string.IsNullOrEmpty(config.ApiKey) ||
                     !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OPENROUTER_API_KEY")) ||
                     !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("JEV_API_KEY"));
        return new
        {
            hasSavedKey = hasKey,
            model = config.ResolvedModel,
            provider = config.ProviderLabel,
            rules = RuleStore.Load(),
            playbooks = Playbooks.All,
            accounts = AccountStore.Load().Select(AccountDto).ToList(),
            jobs = AccountStore.Load().ToDictionary(a => a.Id, a => JobDto(new JobStore(a.Id).Load())),
            defaultAccount = DefaultAccountId(),
            secretBackend = SecretStore.Backend,
            hasClientId = !string.IsNullOrWhiteSpace(config.ClientId),
            dashboardUrl = $"http://127.0.0.1:{_httpPort}/",
            addinManifestUrl = _httpsPort is { } hp ? $"https://localhost:{hp}/manifest.xml" : null,
        };
    }

    private async Task<object> TestKeyAsync(string apiKey, bool saveKey, CancellationToken ct)
    {
        var config = AppConfig.Load();
        var key = config.ResolveApiKey(apiKey);
        var rules = RuleStore.Load();
        var payload = JevPayloadBuilder.Serialize(JevPayloadBuilder.Build(JevPayloadBuilder.SampleMetadata(), rules, "metadata", string.Empty, config.ResolvedModel));
        var parsed = await new JevClient(_http, config.ResolvedEndpoint, key).DecideAsync(payload, JevPayloadBuilder.EstimateCost(payload), rules, ct);
        if (!parsed.Ok) throw new InvalidOperationException(parsed.Error);
        if (saveKey && !string.IsNullOrWhiteSpace(apiKey))
        {
            config.ApiKey = apiKey.Trim();
            config.Save();
        }
        return new
        {
            ok = true,
            model = parsed.Model.Length > 0 ? parsed.Model : config.ResolvedModel,
            provider = parsed.Provider,
            label = RuleValidator.NameById(rules, parsed.RuleId),
            confidence = parsed.Confidence,
            costUsd = parsed.CostUsd,
        };
    }

    private static object ClearKey()
    {
        var config = AppConfig.Load();
        config.ApiKey = null;
        config.Save();
        return new { ok = true };
    }

    private static List<LabelRule> Rules(JsonElement[] args, int index)
    {
        if (args.Length <= index || args[index].ValueKind != JsonValueKind.Array) throw new RuleValidationException("Label rules must be an array.");
        return args[index].Deserialize<List<LabelRule?>>(Json)!.Select(r => r ?? new LabelRule(string.Empty, string.Empty, string.Empty, false)).ToList();
    }

    // ---------------------------------------------------------------------
    // accounts
    // ---------------------------------------------------------------------

    private static object AccountDto(MailAccount a) => new
    {
        id = a.Id,
        email = a.Email,
        kind = a.Kind,
        provider = a.ProviderLabel,
        gmail = a.Gmail,
        host = a.Host,
        port = a.Port,
        ready = MailboxFactory.IsReady(a),
    };

    /// <summary>
    /// The dashboard always names the account. The add-in task pane does not (it runs
    /// inside Outlook), so an empty id falls back to the first ready Microsoft mailbox.
    /// </summary>
    private static MailAccount Account(string idOrEmail)
    {
        if (!string.IsNullOrWhiteSpace(idOrEmail)) return AccountStore.Require(idOrEmail);
        var accounts = AccountStore.Load();
        return accounts.FirstOrDefault(a => a.IsGraph && MailboxFactory.IsReady(a))
            ?? accounts.FirstOrDefault(a => a.IsGraph)
            ?? AccountStore.Resolve(null);
    }

    /// <summary>The account an empty id resolves to (the add-in pane's mailbox), or null when none is configured or it is ambiguous.</summary>
    private static string? DefaultAccountId()
    {
        try { return Account(string.Empty).Id; }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException) { return null; }
    }

    /// <summary>
    /// Add a mailbox. IMAP/Gmail: the password is stored in the secret store and the
    /// connection is verified before the account is kept. Microsoft 365: the account
    /// is created and the device-code sign-in starts (poll authStatus for the code).
    /// </summary>
    private async Task<object?> AddAccountAsync(JsonElement o, CancellationToken ct)
    {
        var email = (Prop(o, "email") ?? string.Empty).Trim();
        if (email.Length < 3 || !email.Contains('@')) throw new ArgumentException("Enter the mailbox address.");
        var kind = (Prop(o, "kind") ?? string.Empty).Trim().ToLowerInvariant();
        var account = new MailAccount { Id = MailAccount.IdFor(email), Email = email };
        switch (kind)
        {
            case "m365":
            case "graph":
                account.Kind = AccountKind.Graph;
                var tenant = (Prop(o, "tenantId") ?? string.Empty).Trim();
                if (tenant.Length > 0) account.TenantId = tenant;
                AccountStore.Add(account);
                return new { ok = true, account = AccountDto(account), signIn = await StartSignInAsync(account.Id) };
            case "gmail":
                account.Kind = AccountKind.Imap;
                account.Gmail = true;
                account.Host = "imap.gmail.com";
                account.Port = 993;
                break;
            case "imap":
                account.Kind = AccountKind.Imap;
                account.Host = (Prop(o, "host") ?? string.Empty).Trim();
                if (account.Host.Length == 0) throw new ArgumentException("Enter the IMAP server host name.");
                account.Port = (int)Num(o, "port", 993);
                if (account.Port is < 1 or > 65535) throw new ArgumentException("The IMAP port must be between 1 and 65535.");
                var user = (Prop(o, "username") ?? string.Empty).Trim();
                if (user.Length > 0 && !string.Equals(user, email, StringComparison.OrdinalIgnoreCase)) account.Username = user;
                break;
            default:
                throw new ArgumentException("Choose the mailbox type: Microsoft 365, Gmail or IMAP.");
        }

        var password = Prop(o, "password") ?? string.Empty;
        if (password.Length == 0) throw new ArgumentException(account.Gmail ? "Enter the Gmail app password." : "Enter the mailbox password.");
        // Verify before persisting anything: a wrong password must not leave a half-configured account.
        var probe = new Imap.ImapMailbox(account, password);
        var identity = await probe.GetIdentityAsync(ct);
        await probe.ConnectAsync(ct);
        AccountStore.Add(account);
        SecretStore.Set(account.Id, password);
        return new { ok = true, account = AccountDto(account), identity };
    }

    private object RemoveAccount(string accountId)
    {
        var account = AccountStore.Require(accountId);
        var job = new JobStore(account.Id).Load();
        if (job is { Status: JobStatus.Running }) throw new InvalidOperationException("Stop processing on this mailbox before removing it.");
        _engines.Remove(account.Id);
        AccountStore.Remove(account.Id);
        return new { ok = true, accounts = AccountStore.Load().Select(AccountDto).ToList() };
    }

    private async Task<object?> TestAccountAsync(string accountId, CancellationToken ct)
    {
        var account = Account(accountId);
        var mailbox = await MailboxFactory.OpenAsync(account, AppConfig.Load(), _http, ct);
        var identity = await mailbox.GetIdentityAsync(ct);
        var caps = await mailbox.ConnectAsync(ct);
        var unread = await mailbox.EstimateAsync("inbox", true, ct);
        var total = await mailbox.EstimateAsync("inbox", false, ct);
        return new { ok = true, identity, provider = caps.ProviderName, labelNoun = caps.LabelNounPlural, canArchive = caps.CanArchive, supportsAllScope = caps.SupportsAllScope, archive = caps.ArchiveDescription, inboxUnread = unread, inboxTotal = total };
    }

    // ---------------------------------------------------------------------
    // Microsoft sign-in (device code) / IMAP readiness
    // ---------------------------------------------------------------------

    private async Task<object> AuthStatusAsync(string accountId, CancellationToken ct)
    {
        var account = Account(accountId);
        if (account.IsImap)
        {
            return SecretStore.Has(account.Id)
                ? new { signedIn = true, account = account.Email }
                : new { signedIn = false, error = "No password is stored for this mailbox. Remove it and add it again." };
        }

        SignInState? state;
        lock (_authGate)
        {
            _signIns.TryGetValue(account.Id, out state);
            if (state?.Task is { IsCompleted: false })
                return new { signedIn = false, pending = true, code = state.Code ?? "…", url = state.Url ?? "https://login.microsoft.com/device" };
        }
        if (!GraphTokenProvider.HasRecord(account.Id)) return new { signedIn = false, error = state?.Error };
        try
        {
            var mailbox = await MailboxFactory.OpenAsync(account, AppConfig.Load(), _http, ct);
            return new { signedIn = true, account = await mailbox.GetIdentityAsync(ct) };
        }
        catch (Exception ex)
        {
            return new { signedIn = false, error = ex.Message };
        }
    }

    private Task<object> StartSignInAsync(string accountId)
    {
        var account = Account(accountId);
        if (!account.IsGraph) throw new InvalidOperationException("Only Microsoft 365 mailboxes use the Microsoft sign-in.");
        var config = AppConfig.Load();
        if (string.IsNullOrWhiteSpace(config.ClientId))
            throw new InvalidOperationException("No Entra ID application (client) id is configured. Run: jevoutlook config set client-id <guid>");

        lock (_authGate)
        {
            if (_signIns.TryGetValue(account.Id, out var running) && running.Task is { IsCompleted: false })
                return Task.FromResult<object>(new { signedIn = false, pending = true, code = running.Code ?? "…", url = running.Url ?? "https://login.microsoft.com/device" });

            var state = new SignInState();
            _signIns[account.Id] = state;
            var ready = new TaskCompletionSource();
            state.Task = Task.Run(async () =>
            {
                try
                {
                    using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(15));
                    await MailboxFactory.SignInGraphAsync(account, config, _http, true, info =>
                    {
                        lock (_authGate) { state.Code = info.UserCode; state.Url = info.VerificationUri.ToString(); }
                        ready.TrySetResult();
                    }, timeout.Token);
                    var accounts = AccountStore.Load();
                    var saved = accounts.FirstOrDefault(a => a.Id == account.Id);
                    if (saved is not null && saved.Email != account.Email) { saved.Email = account.Email; AccountStore.Save(accounts); }
                    _engines.Remove(account.Id);
                }
                catch (Exception ex)
                {
                    lock (_authGate) { state.Error = "Microsoft sign-in failed: " + ex.Message; }
                    ready.TrySetResult();
                }
                finally
                {
                    lock (_authGate) { state.Code = null; state.Url = null; }
                }
            });
            // Wait briefly for the device code to be issued so the page can show it right away.
            return ready.Task.WaitAsync(TimeSpan.FromSeconds(20)).ContinueWith<object>(_ =>
            {
                lock (_authGate)
                {
                    return state.Error is not null
                        ? new { signedIn = false, error = state.Error }
                        : new { signedIn = false, pending = true, code = state.Code ?? "…", url = state.Url ?? "https://login.microsoft.com/device" };
                }
            });
        }
    }

    private object SignOut(string accountId)
    {
        var account = Account(accountId);
        if (account.IsGraph) GraphTokenProvider.DeleteRecord(account.Id); else SecretStore.Delete(account.Id);
        _engines.Remove(account.Id);
        return new { signedIn = false };
    }

    // ---------------------------------------------------------------------
    // processing session (one engine per account)
    // ---------------------------------------------------------------------

    private readonly Dictionary<string, (TriageEngine Engine, string Key)> _engines = new(StringComparer.OrdinalIgnoreCase);

    private async Task<TriageEngine> EngineAsync(MailAccount account, string? apiKey, CollectingSink sink, CancellationToken ct)
    {
        var config = AppConfig.Load();
        var key = config.ResolveApiKey(apiKey);
        if (!_engines.TryGetValue(account.Id, out var entry) || entry.Key != key)
        {
            var mailbox = await MailboxFactory.OpenAsync(account, config, _http, ct);
            entry = (new TriageEngine(mailbox, new JobStore(account.Id), new JevClient(_http, config.ResolvedEndpoint, key), config.ResolvedModel, sink), key);
            _engines[account.Id] = entry;
        }
        entry.Engine.Sink = sink;
        return entry.Engine;
    }

    private async Task<object?> StartJobAsync(JsonElement options, CancellationToken ct)
    {
        var o = options.ValueKind == JsonValueKind.Object ? options : default;
        var account = Account(Prop(o, "account") ?? string.Empty);
        var apiKey = Prop(o, "apiKey");
        var saveKey = o.ValueKind == JsonValueKind.Object && o.TryGetProperty("saveKey", out var sk) && sk.ValueKind == JsonValueKind.True;
        if (saveKey && !string.IsNullOrWhiteSpace(apiKey))
        {
            var config = AppConfig.Load();
            config.ApiKey = apiKey.Trim();
            config.Save();
        }

        var runOptions = RunOptions.Normalize(
            scope: Prop(o, "scope"),
            unreadOnly: !(o.ValueKind == JsonValueKind.Object && o.TryGetProperty("unreadOnly", out var u) && u.ValueKind == JsonValueKind.False),
            dryRun: !(o.ValueKind == JsonValueKind.Object && o.TryGetProperty("dryRun", out var d) && d.ValueKind == JsonValueKind.False),
            mode: Prop(o, "mode"),
            limitRaw: Prop(o, "limit"),
            maxSpendUsd: Num(o, "maxSpendUsd", AppConstants.DefaultMaxSpendUsd),
            metadataThreshold: Num(o, "metadataThreshold", AppConstants.DefaultMetadataThreshold),
            archiveThreshold: Num(o, "archiveThreshold", AppConstants.DefaultArchiveThreshold));

        var rules = o.ValueKind == JsonValueKind.Object && o.TryGetProperty("rules", out var r) && r.ValueKind == JsonValueKind.Array
            ? Rules([r], 0)
            : RuleStore.Load();

        var sink = new CollectingSink();
        var engine = await EngineAsync(account, apiKey, sink, ct);
        var job = await engine.StartAsync(runOptions, rules, ct);
        return new { job = JobDto(job), events = sink.Events, results = sink.Results };
    }

    private async Task<object?> ProcessNextBatchAsync(string jobId, string apiKey, string accountId, CancellationToken ct)
    {
        var account = Account(accountId);
        var job = new JobStore(account.Id).Load();
        if (job is null || job.Id != jobId) throw new InvalidOperationException("This processing session is no longer available.");
        var sink = new CollectingSink();
        if (job.IsRunning)
        {
            var engine = await EngineAsync(account, apiKey, sink, ct);
            await engine.ProcessBatchAsync(job, ct);
        }
        return new { job = JobDto(job), events = sink.Events, results = sink.Results };
    }

    private async Task<object?> ResumeJobAsync(string accountId, CancellationToken ct)
    {
        var account = Account(accountId);
        var engine = await EngineAsync(account, null, new CollectingSink(), ct);
        return JobDto(engine.Resume(null));
    }

    // ---------------------------------------------------------------------
    // single message (task pane)
    // ---------------------------------------------------------------------

    private async Task<IMailbox> MailboxAsync(string accountId, CancellationToken ct) =>
        await MailboxFactory.OpenAsync(Account(accountId), AppConfig.Load(), _http, ct);

    /// <summary>
    /// Classify one message with the same two-stage policy as a run (metadata first,
    /// body when confidence is below the metadata threshold). Nothing is written.
    /// The pseudo id <c>latest</c> selects the newest Inbox message (testing outside Outlook).
    /// </summary>
    private async Task<object?> ClassifyItemAsync(string itemId, string apiKey, string accountId, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(itemId)) throw new ArgumentException("No message id was provided.");
        var config = AppConfig.Load();
        var jev = new JevClient(_http, config.ResolvedEndpoint, config.ResolveApiKey(apiKey));
        var rules = RuleStore.Load();
        var mailbox = await MailboxAsync(accountId, ct);
        await mailbox.ConnectAsync(ct);

        if (itemId == "latest")
        {
            var page = await mailbox.ListMessagesAsync("inbox", false, await mailbox.GetInitialCursorAsync("inbox", ct), false, 1, ct);
            itemId = page.Items.FirstOrDefault()?.Id ?? throw new InvalidOperationException("The Inbox is empty.");
        }

        var metaParts = await mailbox.ReadMessagesAsync([itemId], ReadMode.Metadata, ct);
        if (!metaParts[0].Ok || metaParts[0].Metadata is null)
        {
            throw new InvalidOperationException(metaParts[0].NotFound
                ? "This message could not be found in the mailbox (it may have been moved or deleted)."
                : "The mailbox could not read the message: " + metaParts[0].Error);
        }
        var metadata = metaParts[0].Metadata!;
        var stages = new List<object>();
        double cost = 0;

        JevResult Check(JevResult r)
        {
            cost += r.CostUsd;
            if (!r.Ok) throw new InvalidOperationException(r.Error);
            return r;
        }
        object StageDto(string stage, JevResult r) => new
        {
            stage,
            ruleId = r.RuleId,
            label = RuleValidator.NameById(rules, r.RuleId),
            confidence = r.Confidence,
            probabilities = (r.Probabilities ?? new Dictionary<string, double>())
                .Select(kv => new { label = rules[int.Parse(kv.Key[1..], CultureInfo.InvariantCulture)].Name, p = kv.Value })
                .OrderByDescending(x => x.p).ToList(),
        };

        var metaPayload = JevPayloadBuilder.Serialize(JevPayloadBuilder.Build(metadata, rules, "metadata", string.Empty, config.ResolvedModel));
        var first = Check(await jev.DecideAsync(metaPayload, JevPayloadBuilder.EstimateCost(metaPayload), rules, ct));
        stages.Add(StageDto("metadata", first));
        var final = first;
        var finalStage = "metadata";

        if (first.Confidence < AppConstants.DefaultMetadataThreshold)
        {
            var fullParts = await mailbox.ReadMessagesAsync([itemId], ReadMode.Full, ct);
            if (fullParts[0].Ok)
            {
                var body = MessageContent.Compact(fullParts[0].FullText);
                if (body.Length > 0)
                {
                    var fullPayload = JevPayloadBuilder.Serialize(JevPayloadBuilder.Build(metadata, rules, "full", body, config.ResolvedModel));
                    final = Check(await jev.DecideAsync(fullPayload, JevPayloadBuilder.EstimateCost(fullPayload), rules, ct));
                    finalStage = "full";
                    stages.Add(StageDto("full", final));
                }
            }
        }

        var rule = RuleValidator.ById(rules, final.RuleId)!;
        return new
        {
            id = itemId,
            subject = metadata.Subject,
            from = metadata.From,
            receivedDateTime = metadata.Date,
            categories = metadata.Categories,
            alreadyTriaged = TriageEngine.HasConfiguredCategory(metadata.Categories, rules),
            stages,
            final = new { ruleId = rule.Id, label = rule.Name, spam = rule.Spam, confidence = final.Confidence, stage = finalStage, probabilities = ((dynamic)stages[^1]).probabilities },
            costUsd = cost,
        };
    }

    /// <summary>Apply one decision to one message: merge the label, optionally archive.</summary>
    private async Task<object?> ApplyItemAsync(string itemId, string ruleId, bool archive, string accountId, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(itemId) || itemId == "latest") throw new ArgumentException("Classify the message first.");
        var rules = RuleStore.Load();
        var rule = RuleValidator.ById(rules, ruleId) ?? throw new ArgumentException("Unknown category.");
        if (archive && !rule.Spam) throw new InvalidOperationException($"“{rule.Name}” is not archive eligible.");
        var mailbox = await MailboxAsync(accountId, ct);
        var caps = await mailbox.ConnectAsync(ct);
        if (archive && !caps.CanArchive) throw new InvalidOperationException("This mailbox has no archive destination.");

        var names = await mailbox.EnsureLabelsAsync(rules.Select(r => r.Name).ToList(), ct);

        var current = await mailbox.ReadMessagesAsync([itemId], ReadMode.Labels, ct);
        if (!current[0].Ok) throw new InvalidOperationException(current[0].NotFound ? "This message could not be found in the mailbox." : "The mailbox could not read the message: " + current[0].Error);
        var merged = new List<string>(current[0].Labels);
        var categoryName = names.GetValueOrDefault(rule.Name.ToLowerInvariant(), rule.Name);
        if (!merged.Any(c => string.Equals(c, categoryName, StringComparison.OrdinalIgnoreCase))) merged.Add(categoryName);

        var outcome = await mailbox.ApplyLabelsAsync([(itemId, merged)], ct);
        if (!outcome.Ok) throw new InvalidOperationException(outcome.Error);
        var archived = false;
        if (archive)
        {
            var move = await mailbox.ArchiveAsync([itemId], ct);
            if (!move.Ok) throw new InvalidOperationException(move.Error);
            archived = true;
        }
        return new { ok = true, categories = merged, archived };
    }

    private static object CancelJob(string accountId)
    {
        var store = new JobStore(Account(accountId).Id);
        var job = store.Load();
        if (job is null) return new { ok = true };
        if (job.Status is JobStatus.Completed or JobStatus.Cancelled) return new { ok = true, job = JobDto(job) };
        job.Status = JobStatus.Cancelled;
        job.StopReason = "user-cancelled";
        store.Save(job);
        return new { ok = true, job = JobDto(job) };
    }

    private static object ClearJob(string accountId)
    {
        var store = new JobStore(Account(accountId).Id);
        var job = store.Load();
        if (job is { Status: JobStatus.Running }) throw new InvalidOperationException("Stop processing before clearing the session.");
        store.Delete();
        return new { ok = true };
    }

    /// <summary>The subset of the job the dashboard renders (no pending ids, no skipped-id list).</summary>
    private static object? JobDto(TriageJob? job)
    {
        if (job is null) return null;
        return new
        {
            id = job.Id,
            status = job.Status,
            stopReason = job.StopReason,
            lastError = job.LastError,
            dryRun = job.DryRun,
            mode = job.Mode,
            scope = job.Scope,
            unreadOnly = job.UnreadOnly,
            limit = job.Limit is { } l ? l.ToString(CultureInfo.InvariantCulture) : "all",
            maxSpendUsd = job.MaxSpendUsd,
            metadataThreshold = job.MetadataThreshold,
            archiveThreshold = job.ArchiveThreshold,
            initialEstimate = job.InitialEstimate,
            target = job.Target,
            processed = job.Processed,
            metadataOnly = job.MetadataOnly,
            fullBody = job.FullBody,
            archived = job.Archived,
            failed = job.Failed,
            skipped = job.Skipped,
            providerRetries = job.ProviderRetries,
            graphRateLimitRetries = job.GraphRateLimitRetries,
            modelResponseSkips = job.ModelResponseSkips,
            modelRequests = job.ModelRequests,
            spentUsd = job.SpentUsd,
            reportedCostUsd = job.ReportedCostUsd,
            tokenCalculatedCostUsd = job.TokenCalculatedCostUsd,
            estimatedCostUsd = job.EstimatedCostUsd,
            ruleLabels = job.RuleLabels,
            labelCounts = job.LabelCounts,
            elapsedMs = job.ElapsedNowMs(),
            elapsedMeasuredAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
            createdAt = job.CreatedAt,
            finishedAt = job.FinishedAt,
            signedInAs = job.SignedInAs,
            accountId = job.AccountId,
            provider = job.Provider,
        };
    }

    // ---------------------------------------------------------------------
    // helpers
    // ---------------------------------------------------------------------

    private static string Str(JsonElement[] args, int index) =>
        args.Length > index && args[index].ValueKind == JsonValueKind.String ? args[index].GetString() ?? string.Empty : string.Empty;

    private static bool Bool(JsonElement[] args, int index) => args.Length > index && args[index].ValueKind == JsonValueKind.True;

    private static string? Prop(JsonElement o, string name) =>
        o.ValueKind == JsonValueKind.Object && o.TryGetProperty(name, out var v)
            ? v.ValueKind switch { JsonValueKind.String => v.GetString(), JsonValueKind.Number => v.GetRawText(), _ => null }
            : null;

    private static double Num(JsonElement o, string name, double fallback)
    {
        if (o.ValueKind != JsonValueKind.Object || !o.TryGetProperty(name, out var v)) return fallback;
        return v.ValueKind switch
        {
            JsonValueKind.Number when v.TryGetDouble(out var d) => d,
            JsonValueKind.String when double.TryParse(v.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var p) => p,
            _ => double.NaN,
        };
    }

    /// <summary>Buffers one request's events and result rows for the JSON response.</summary>
    private sealed class CollectingSink : ITriageSink
    {
        public List<object> Events { get; } = [];
        public List<ResultRow> Results { get; } = [];
        public void Event(string level, string message) => Events.Add(new { level, message });
        public void Result(ResultRow row) => Results.Add(row);
        public void BatchCompleted(TriageJob job) { }
    }
}
