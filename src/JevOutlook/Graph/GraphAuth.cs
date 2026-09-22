using Azure.Core;
using Azure.Identity;
using JevOutlook.Storage;

namespace JevOutlook.Graph;

/// <summary>
/// Delegated Microsoft Graph sign-in through Azure.Identity. The token cache
/// is persisted by MSAL (Keychain on macOS, DPAPI on Windows, libsecret on
/// Linux) and the <see cref="AuthenticationRecord"/> is saved on disk so
/// later runs authenticate silently.
/// </summary>
public sealed class GraphTokenProvider
{
    private readonly TokenCredential _credential;
    private readonly TokenRequestContext _context = new(AppConstants.GraphScopes);

    public GraphTokenProvider(AppConfig config, AuthenticationRecord? record, bool deviceCode)
    {
        if (string.IsNullOrWhiteSpace(config.ClientId))
        {
            throw new InvalidOperationException(
                "No Entra ID application (client) id is configured. Run: jevoutlook config set client-id <guid>");
        }

        var cache = new TokenCachePersistenceOptions { Name = AppPaths.TokenCacheName };

        if (deviceCode || config.DeviceCode)
        {
            _credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
            {
                ClientId = config.ClientId,
                TenantId = config.TenantId,
                TokenCachePersistenceOptions = cache,
                AuthenticationRecord = record,
                // Never prompt in the middle of a batch: GetTokenAsync throws AuthenticationRequiredException
                // instead, which is turned into "run: jevoutlook auth". SignInAsync prompts explicitly.
                DisableAutomaticAuthentication = true,
                DeviceCodeCallback = (info, _) =>
                {
                    Console.WriteLine(info.Message);
                    return Task.CompletedTask;
                },
            });
        }
        else
        {
            _credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                ClientId = config.ClientId,
                TenantId = config.TenantId,
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
            throw new InvalidOperationException("Microsoft sign-in is required or has expired. Run: jevoutlook auth");
        }
        catch (CredentialUnavailableException ex)
        {
            throw new InvalidOperationException("Microsoft sign-in is unavailable: " + ex.Message + " Run: jevoutlook auth");
        }
    }

    // ----- Authentication record persistence -------------------------------

    public static async Task SaveRecordAsync(AuthenticationRecord record, CancellationToken ct)
    {
        JsonStore.EnsureDirectory(AppPaths.Root);
        await using var stream = new MemoryStream();
        await record.SerializeAsync(stream, ct);
        File.WriteAllBytes(AppPaths.AuthRecord, stream.ToArray());
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(AppPaths.AuthRecord, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        }
    }

    public static async Task<AuthenticationRecord?> LoadRecordAsync(CancellationToken ct)
    {
        if (!File.Exists(AppPaths.AuthRecord)) return null;
        try
        {
            await using var stream = File.OpenRead(AppPaths.AuthRecord);
            return await AuthenticationRecord.DeserializeAsync(stream, ct);
        }
        catch (Exception)
        {
            return null;
        }
    }

    public static void DeleteRecord() => JsonStore.Delete(AppPaths.AuthRecord);
}
