using Azure.Identity;
using JevOutlook.Graph;
using JevOutlook.Imap;
using JevOutlook.Storage;

namespace JevOutlook.Mail;

/// <summary>Builds the right <see cref="IMailbox"/> for a configured account.</summary>
public static class MailboxFactory
{
    /// <summary>
    /// Open a mailbox for processing. Graph accounts must already hold a saved
    /// sign-in record; IMAP accounts must have a stored password.
    /// </summary>
    public static async Task<IMailbox> OpenAsync(MailAccount account, AppConfig config, HttpClient http, CancellationToken ct)
    {
        if (account.IsGraph)
        {
            var record = await GraphTokenProvider.LoadRecordAsync(account.Id, ct)
                ?? throw new InvalidOperationException($"Account '{account.Id}' is not signed in to Microsoft. Run: jevoutlook account login {account.Id}");
            return new GraphMailClient(http, new GraphTokenProvider(config, account, record, config.DeviceCode));
        }
        if (account.IsImap)
        {
            var password = SecretStore.Require(account.Id);
            return new ImapMailbox(account, password);
        }
        throw new InvalidOperationException($"Account '{account.Id}' has an unknown kind '{account.Kind}'.");
    }

    /// <summary>Interactive Microsoft sign-in for a Graph account; saves the record for silent reuse.</summary>
    public static async Task<IMailbox> SignInGraphAsync(MailAccount account, AppConfig config, HttpClient http, bool deviceCode,
        Action<DeviceCodeInfo>? deviceCodeCallback, CancellationToken ct)
    {
        var provider = new GraphTokenProvider(config, account, null, deviceCode, deviceCodeCallback);
        var record = await provider.SignInAsync(ct);
        await GraphTokenProvider.SaveRecordAsync(account.Id, record, ct);
        if (string.IsNullOrWhiteSpace(account.Email) || account.Email == "microsoft-account")
        {
            account.Email = record.Username ?? account.Email;
        }
        return new GraphMailClient(http, provider);
    }

    /// <summary>True when the account holds what it needs to open without prompting.</summary>
    public static bool IsReady(MailAccount account) =>
        account.IsGraph ? GraphTokenProvider.HasRecord(account.Id) : account.IsImap && SecretStore.Has(account.Id);
}
