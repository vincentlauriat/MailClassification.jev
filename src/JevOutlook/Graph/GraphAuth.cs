using Azure.Core;
using Azure.Identity;
using JevOutlook.Storage;

namespace JevOutlook.Graph;

/// <summary>
/// Delegated Microsoft Graph sign-in through Azure.Identity, one record per
/// account. The token cache is persisted by MSAL (Keychain on macOS, DPAPI on
/// Windows, libsecret on Linux) and the <see cref="AuthenticationRecord"/> is
/// saved in the account directory so later runs authenticate silently.
/// </summary>
public sealed class GraphTokenProvider
{
    private readonly TokenCredential _credential;
    private readonly TokenRequestContext _context = new(AppConstants.GraphScopes);

    public GraphTokenProvider(AppConfig config, MailAccount account, AuthenticationRecord? record, bool deviceCode, Action<DeviceCodeInfo>? deviceCodeCallback = null)
    {
        if (string.IsNullOrWhiteSpace(config.ClientId))
        {
            throw new InvalidOperationException(
                "No Entra ID application (client) id is configured. Run: jevoutlook config set client-id <guid>");
        }

        var cache = new TokenCachePersistenceOptions { Name = AppPaths.TokenCacheName };
        var tenant = string.IsNullOrWhiteSpace(account.TenantId) ? config.TenantId : account.TenantId;

        if (deviceCode || config.DeviceCode)
        {
            _credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
            {
                ClientId = config.ClientId,
                TenantId = tenant,
                TokenCachePersistenceOptions = cache,
                AuthenticationRecord = record,
                // Never prompt in the middle of a batch: GetTokenAsync throws AuthenticationRequiredException
                // instead, which is turned into "run: jevoutlook account login". SignInAsync prompts explicitly.
                DisableAutomaticAuthentication = true,
                DeviceCodeCallback = (info, _) =>
                {
                    if (deviceCodeCallback is not null) deviceCodeCallback(info);
                    else Console.WriteLine(info.Message);
                    return Task.CompletedTask;
                },
            });
        }
        else
        {
            _credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                ClientId = config.ClientId,
                TenantId = tenant,
                RedirectUri = new Uri("http://localhost"),
                TokenCachePersistenceOptions = cache,
                AuthenticationRecord = record,
                DisableAutomaticAuthentication = true,
            });
        }
    }

    /// <summary>Interactive sign-in. Returns the record to persist for silent reuse.</summary>
    public async Task<AuthenticationRecord> SignInAsync(CancellationToken ct)
    {
        return _credential switch
        {
            InteractiveBrowserCredential browser => await browser.AuthenticateAsync(_context, ct),
            DeviceCodeCredential device => await device.AuthenticateAsync(_context, ct),
            _ => throw new InvalidOperationException("Unsupported credential."),
        };
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken ct)
    {
        try
        {
            var token = await _credential.GetTokenAsync(_context, ct);
            return token.Token;
        }
        catch (AuthenticationRequiredException)
        {
            throw new InvalidOperationException("Microsoft sign-in is required or has expired. Run: jevoutlook account login <id>");
        }
        catch (CredentialUnavailableException ex)
        {
            throw new InvalidOperationException("Microsoft sign-in is unavailable: " + ex.Message + " Run: jevoutlook account login <id>");
        }
    }

    // ----- Authentication record persistence (per account) ------------------

    public static async Task SaveRecordAsync(string accountId, AuthenticationRecord record, CancellationToken ct)
    {
        var path = AppPaths.AccountAuthRecord(accountId);
        JsonStore.EnsureDirectory(Path.GetDirectoryName(path)!);
        await using var stream = new MemoryStream();
        await record.SerializeAsync(stream, ct);
        File.WriteAllBytes(path, stream.ToArray());
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        }
    }

    public static Task<AuthenticationRecord?> LoadRecordAsync(string accountId, CancellationToken ct) =>
        LoadRecordFromAsync(AppPaths.AccountAuthRecord(accountId), ct);

    public static async Task<AuthenticationRecord?> LoadRecordFromAsync(string path, CancellationToken ct)
    {
        if (!File.Exists(path)) return null;
        try
        {
            await using var stream = File.OpenRead(path);
            return await AuthenticationRecord.DeserializeAsync(stream, ct);
        }
        catch (Exception)
        {
            return null;
        }
    }

    public static bool HasRecord(string accountId) => File.Exists(AppPaths.AccountAuthRecord(accountId));

    public static void DeleteRecord(string accountId) => JsonStore.Delete(AppPaths.AccountAuthRecord(accountId));

    /// <summary>
    /// Move the pre-multi-account state (one Microsoft mailbox at the root of the
    /// state directory) into an account entry. Runs once; a no-op afterwards.
    /// </summary>
    public static async Task<MailAccount?> MigrateLegacyAsync(CancellationToken ct)
    {
        if (File.Exists(AppPaths.Accounts) || !File.Exists(AppPaths.LegacyAuthRecord)) return null;
        var record = await LoadRecordFromAsync(AppPaths.LegacyAuthRecord, ct);
        if (record is null) return null;

        var email = record.Username ?? "microsoft-account";
        var account = new MailAccount { Id = MailAccount.IdFor(email), Email = email, Kind = AccountKind.Graph };
        AccountStore.Add(account);
        File.Move(AppPaths.LegacyAuthRecord, AppPaths.AccountAuthRecord(account.Id), overwrite: true);
        // A checkpointed session from the old layout cannot be resumed (its cursor type changed): drop it.
        JsonStore.Delete(AppPaths.LegacyJob);
        JsonStore.Delete(AppPaths.LegacyJobRules);
        return account;
    }
}
