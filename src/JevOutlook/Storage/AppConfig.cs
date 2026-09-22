namespace JevOutlook.Storage;

/// <summary>
/// Persistent configuration: the Entra ID app registration used to sign in,
/// and the model provider. The API key is optional here — the environment
/// variables <c>OPENROUTER_API_KEY</c> / <c>JEV_API_KEY</c> take precedence.
/// </summary>
public sealed class AppConfig
{
    public string? ClientId { get; set; }
    public string TenantId { get; set; } = "common";

    /// <summary><c>openrouter</c> (default) or <c>typesafe</c>.</summary>
    public string Provider { get; set; } = "openrouter";
    public string? EndpointUrl { get; set; }
    public string? Model { get; set; }
    public string? ApiKey { get; set; }
    public bool DeviceCode { get; set; }

    public string ResolvedEndpoint => !string.IsNullOrWhiteSpace(EndpointUrl)
        ? EndpointUrl
        : IsTypeSafe ? AppConstants.TypeSafeUrl : AppConstants.OpenRouterDecisionsUrl;

    public string ResolvedModel => !string.IsNullOrWhiteSpace(Model)
        ? Model
        : IsTypeSafe ? AppConstants.TypeSafeModel : AppConstants.OpenRouterModel;

    public bool IsTypeSafe => string.Equals(Provider, "typesafe", StringComparison.OrdinalIgnoreCase);

    public string ProviderLabel => IsTypeSafe ? "TypeSafe" : "OpenRouter";

    public static AppConfig Load() => JsonStore.Load<AppConfig>(AppPaths.Config) ?? new AppConfig();

    public void Save() => JsonStore.Save(AppPaths.Config, this);

    /// <summary>Key can be supplied on the command line, from the environment, or from the config file.</summary>
    public string ResolveApiKey(string? explicitKey)
    {
        var key = (explicitKey ?? string.Empty).Trim();
        if (key.Length == 0) key = (Environment.GetEnvironmentVariable("JEV_API_KEY") ?? string.Empty).Trim();
        if (key.Length == 0 && !IsTypeSafe) key = (Environment.GetEnvironmentVariable("OPENROUTER_API_KEY") ?? string.Empty).Trim();
        if (key.Length == 0) key = (ApiKey ?? string.Empty).Trim();
        if (key.Length == 0)
            throw new InvalidOperationException(
                $"Enter an API key for {ProviderLabel} (jevoutlook key set <key>) or export OPENROUTER_API_KEY / JEV_API_KEY.");
        return key;
    }
}
