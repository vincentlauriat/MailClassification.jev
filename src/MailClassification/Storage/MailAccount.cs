using System.Text.RegularExpressions;

namespace JevOutlook.Storage;

public static class AccountKind
{
    /// <summary>Microsoft 365 / Outlook.com through Microsoft Graph (device-code or browser sign-in).</summary>
    public const string Graph = "graph";
    /// <summary>IMAP with a password (Gmail app password, or any IMAP server).</summary>
    public const string Imap = "imap";
}

/// <summary>One configured mailbox. Secrets are never stored here (see <see cref="SecretStore"/>).</summary>
public sealed class MailAccount
{
    public string Id { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Kind { get; set; } = AccountKind.Graph;
    public string? DisplayName { get; set; }

    // IMAP
    public string? Host { get; set; }
    public int Port { get; set; } = 993;
    /// <summary>Gmail over IMAP: labels via X-GM-LABELS, archive = remove from Inbox.</summary>
    public bool Gmail { get; set; }
    /// <summary>Login name when it differs from <see cref="Email"/>.</summary>
    public string? Username { get; set; }

    // Graph
    /// <summary>Overrides the global tenant id (e.g. "consumers" for a personal account).</summary>
    public string? TenantId { get; set; }

    public long CreatedAt { get; set; }

    public bool IsGraph => string.Equals(Kind, AccountKind.Graph, StringComparison.OrdinalIgnoreCase);
    public bool IsImap => string.Equals(Kind, AccountKind.Imap, StringComparison.OrdinalIgnoreCase);
    public string LoginName => string.IsNullOrWhiteSpace(Username) ? Email : Username!;
    public string ProviderLabel => IsGraph ? "Microsoft 365" : Gmail ? "Gmail" : "IMAP";

    /// <summary>Stable directory-safe id derived from the address: "alice@contoso.com" → "alice-contoso.com".</summary>
    public static string IdFor(string email)
    {
        var slug = Regex.Replace(email.Trim().ToLowerInvariant(), "[^a-z0-9._-]+", "-").Trim('-', '.');
        return slug.Length == 0 ? "account" : slug;
    }
}

/// <summary>The account registry (<c>accounts.json</c>): a list, first = default.</summary>
public static class AccountStore
{
    public static List<MailAccount> Load() => JsonStore.Load<List<MailAccount>>(AppPaths.Accounts) ?? [];

    public static void Save(List<MailAccount> accounts) => JsonStore.Save(AppPaths.Accounts, accounts);

    /// <summary>Find by id or by e-mail address (case-insensitive).</summary>
    public static MailAccount? Find(string idOrEmail)
    {
        var key = idOrEmail.Trim();
        return Load().FirstOrDefault(a =>
            string.Equals(a.Id, key, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(a.Email, key, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(a.Id, MailAccount.IdFor(key), StringComparison.OrdinalIgnoreCase));
    }

    public static MailAccount Require(string idOrEmail) =>
        Find(idOrEmail) ?? throw new ArgumentException($"Unknown account '{idOrEmail}'. Run 'jevoutlook account list'.");

    /// <summary>
    /// The account to use when none is named: the only one configured, or an error
    /// listing the choices. Never guesses between several mailboxes.
    /// </summary>
    public static MailAccount Resolve(string? idOrEmail)
    {
        if (!string.IsNullOrWhiteSpace(idOrEmail)) return Require(idOrEmail);
        var accounts = Load();
        return accounts.Count switch
        {
            0 => throw new InvalidOperationException("No mailbox is configured. Add one with: jevoutlook account add <email> --m365 | --gmail | --imap <host>"),
            1 => accounts[0],
            _ => throw new ArgumentException("Several mailboxes are configured; choose one with --account <id>: " + string.Join(", ", accounts.Select(a => a.Id))),
        };
    }

    public static MailAccount Add(MailAccount account)
    {
        var accounts = Load();
        if (accounts.Any(a => string.Equals(a.Id, account.Id, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"An account with id '{account.Id}' already exists. Remove it first: jevoutlook account remove {account.Id}");
        account.CreatedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        accounts.Add(account);
        Save(accounts);
        JsonStore.EnsureDirectory(AppPaths.AccountDir(account.Id));
        return account;
    }

    public static bool Remove(string id)
    {
        var accounts = Load();
        var removed = accounts.RemoveAll(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase));
        if (removed == 0) return false;
        Save(accounts);
        SecretStore.Delete(id);
        var dir = AppPaths.AccountDir(id);
        if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
        return true;
    }
}
