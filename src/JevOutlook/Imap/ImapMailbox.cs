using System.Net.Sockets;
using JevOutlook.Mail;
using JevOutlook.Storage;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;

namespace JevOutlook.Imap;

/// <summary>
/// IMAP implementation of <see cref="IMailbox"/> over MailKit.
/// Gmail: labels are Gmail labels (X-GM-LABELS), archive removes the message from
/// the Inbox (it stays in All Mail). Generic servers (Dovecot…): labels are IMAP
/// keywords (custom flags), archive moves the message to an "Archive" folder.
/// One connection per instance; calls are serialised because IMAP is single-stream.
/// </summary>
public sealed class ImapMailbox : IMailbox, IDisposable, IAsyncDisposable
{
    private const int TimeoutMs = 60_000;
    private static readonly string[] SystemLabels = ["\\Inbox"];

    private readonly MailAccount _account;
    private readonly string _password;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private ImapClient? _client;
    private bool _keywordsSupported;
    private bool _connectedOnce;
    private MailboxCapabilities _capabilities;

    public ImapMailbox(MailAccount account, string password)
    {
        _account = account;
        _password = password;
        _capabilities = account.Gmail
            ? new MailboxCapabilities("Gmail", "label", "labels", true, true, "Archive removes the message from the Inbox (it stays in All Mail).")
            : new MailboxCapabilities("IMAP", "keyword", "keywords", true, false, "Archive moves the message to the Archive folder.");
    }

    public MailboxCapabilities Capabilities => _capabilities;

    private bool Gmail => _account.Gmail;

    // =====================================================================
    // Connection plumbing
    // =====================================================================

    private async Task<ImapClient> EnsureClientAsync(CancellationToken ct)
    {
        if (_client is { IsConnected: true, IsAuthenticated: true }) return _client;
        _client?.Dispose();
        var client = new ImapClient { Timeout = TimeoutMs };
        try
        {
            var options = _account.Port == 993 ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
            var host = _account.Host ?? throw new MailboxException(false, $"Account '{_account.Id}' has no IMAP host configured.");
            await client.ConnectAsync(host, _account.Port, options, ct);
            await client.AuthenticateAsync(_account.LoginName, _password, ct);
        }
        catch
        {
            client.Dispose();
            throw;
        }
        _client = client;
        _connectedOnce = true;
        return client;
    }

    /// <summary>
    /// Run one operation on the (single) connection. A dropped connection is
    /// reconnected once and the operation retried. Exceptions are mapped to
    /// <see cref="MailboxException"/> so the engine can pause or fail cleanly.
    /// </summary>
    private async Task<T> WithClientAsync<T>(Func<ImapClient, Task<T>> operation, CancellationToken ct)
    {
        await _gate.WaitAsync(ct);
        try
        {
            for (var attempt = 0; ; attempt++)
            {
                ImapClient client;
                try
                {
                    client = await EnsureClientAsync(ct);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    throw Map(ex);
                }

                try
                {
                    return await operation(client);
                }
                catch (Exception ex) when (attempt == 0 && IsConnectionLoss(ex) && !ct.IsCancellationRequested)
                {
                    _client?.Dispose();
                    _client = null;
                    continue;
                }
                catch (Exception ex) when (ex is not OperationCanceledException and not MailboxException)
                {
                    throw Map(ex);
                }
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    private static bool IsConnectionLoss(Exception ex) =>
        ex is IOException or SocketException or ImapProtocolException or ServiceNotConnectedException or ServiceNotAuthenticatedException;

    private MailboxException Map(Exception ex) => ex switch
    {
        MailboxException m => m,
        AuthenticationException => new MailboxException(false,
            $"{Provider} rejected the password for {_account.Email}. Store a new one with: jevoutlook account password {_account.Id}"),
        ImapCommandException c => new MailboxException(false, $"{Provider} refused a command ({c.Response}): {Clip(c.ResponseText.Length > 0 ? c.ResponseText : c.Message)}"),
        ImapProtocolException or IOException or SocketException or TimeoutException or ServiceNotConnectedException or ServiceNotAuthenticatedException =>
            new MailboxException(true, $"{Provider} connection problem with {_account.Host}: {Clip(ex.Message)}"),
        FolderNotFoundException f => new MailboxException(false, $"{Provider} folder not found: {f.FolderName}"),
        NotSupportedException => new MailboxException(false, $"{Provider} does not support a required IMAP feature: {Clip(ex.Message)}"),
        _ => new MailboxException(false, $"{Provider} error: {Clip(ex.Message)}"),
    };

    private string Provider => _capabilities.ProviderName;

    private static string Clip(string text, int max = 300) => text.Length > max ? text[..max] : text;

    // =====================================================================
    // Folders
    // =====================================================================

    private async Task<IMailFolder> FolderAsync(ImapClient client, string scope, FolderAccess access, CancellationToken ct)
    {
        IMailFolder folder;
        if (scope == ImapSupport.InboxScope)
        {
            folder = client.Inbox;
        }
        else if (scope == ImapSupport.AllScope)
        {
            if (!Gmail) throw new MailboxException(false, "This IMAP mailbox can only be scanned folder by folder: use the Inbox scope.");
            folder = await AllMailAsync(client, ct);
        }
        else
        {
            throw new MailboxException(false, $"Unknown scope '{scope}'.");
        }
        if (!folder.IsOpen || folder.Access < access)
        {
            await folder.OpenAsync(access, ct);
        }
        return folder;
    }

    private static async Task<IMailFolder> AllMailAsync(ImapClient client, CancellationToken ct)
    {
        if (client.Capabilities.HasFlag(ImapCapabilities.SpecialUse) || client.Capabilities.HasFlag(ImapCapabilities.XList))
        {
            try { if (client.GetFolder(SpecialFolder.All) is { } special) return special; }
            catch (Exception ex) when (ex is FolderNotFoundException or NotSupportedException) { }
        }
        foreach (var folder in await client.GetFoldersAsync(client.PersonalNamespaces[0], StatusItems.None, false, ct))
        {
            if (folder.Attributes.HasFlag(FolderAttributes.All)) return folder;
            if (string.Equals(folder.FullName, "[Gmail]/All Mail", StringComparison.OrdinalIgnoreCase)) return folder;
        }
        throw new MailboxException(false, "The Gmail 'All Mail' folder could not be found.");
    }

    private static async Task<IMailFolder> ArchiveFolderAsync(ImapClient client, CancellationToken ct)
    {
        if (client.Capabilities.HasFlag(ImapCapabilities.SpecialUse) || client.Capabilities.HasFlag(ImapCapabilities.XList))
        {
            try { if (client.GetFolder(SpecialFolder.Archive) is { } special) return special; }
            catch (Exception ex) when (ex is FolderNotFoundException or NotSupportedException) { }
        }
        var ns = client.PersonalNamespaces[0];
        foreach (var folder in await client.GetFoldersAsync(ns, StatusItems.None, false, ct))
        {
            if (folder.Attributes.HasFlag(FolderAttributes.Archive)) return folder;
            if (folder.Name.Equals("Archive", StringComparison.OrdinalIgnoreCase) || folder.Name.Equals("Archives", StringComparison.OrdinalIgnoreCase)) return folder;
        }
        var root = client.GetFolder(ns);
        return await root.CreateAsync("Archive", true, ct) ?? throw new MailboxException(false, "The Archive folder could not be created.");
    }

    private static async Task<IMailFolder?> FindLabelFolderAsync(ImapClient client, string label, CancellationToken ct)
    {
        var root = client.GetFolder(client.PersonalNamespaces[0]);
        foreach (var folder in await root.GetSubfoldersAsync(false, ct))
        {
            if (string.Equals(folder.FullName, label, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(folder.Name, label, StringComparison.OrdinalIgnoreCase)) return folder;
        }
        return null;
    }

    // =====================================================================
    // IMailbox
    // =====================================================================

    public Task<string> GetIdentityAsync(CancellationToken ct) =>
        WithClientAsync(_ => Task.FromResult(_account.Email), ct);

    public async Task<MailboxCapabilities> ConnectAsync(CancellationToken ct)
    {
        return await WithClientAsync(async client =>
        {
            var inbox = await FolderAsync(client, ImapSupport.InboxScope, FolderAccess.ReadWrite, ct);
            _keywordsSupported = Gmail || inbox.PermanentFlags.HasFlag(MessageFlags.UserDefined);
            return _capabilities;
        }, ct);
    }

    public async Task<long> EstimateAsync(string scope, bool unreadOnly, CancellationToken ct)
    {
        try
        {
            return await WithClientAsync(async client =>
            {
                var folder = await FolderAsync(client, scope, FolderAccess.ReadWrite, ct);
                return (long)(unreadOnly ? folder.Unread : folder.Count);
            }, ct);
        }
        catch (MailboxException)
        {
            return -1;
        }
    }

    public Task<string> GetInitialCursorAsync(string scope, CancellationToken ct) =>
        WithClientAsync(async client =>
        {
            var folder = await FolderAsync(client, scope, FolderAccess.ReadWrite, ct);
            var next = folder.UidNext?.Id ?? 0u;
            if (next == 0)
            {
                // UIDNEXT not advertised: derive it from the highest existing UID.
                var uids = await folder.SearchAsync(SearchQuery.All, ct);
                next = (uids.Count > 0 ? uids.Max(u => u.Id) : 0u) + 1;
            }
            return ImapSupport.FormatSortKey(folder.UidValidity, next);
        }, ct);

    public Task<ListPage> ListMessagesAsync(string scope, bool unreadOnly, string cursor, bool exclusive, int top, CancellationToken ct) =>
        WithClientAsync(async client =>
        {
            var (validity, uid) = ImapSupport.ParseSortKey(cursor);
            var folder = await FolderAsync(client, scope, FolderAccess.ReadWrite, ct);
            CheckValidity(folder, validity);
            var upper = exclusive ? uid - 1 : uid;
            if (upper < 1) return new ListPage([], false);

            SearchQuery query = SearchQuery.Uids(new UniqueIdRange(new UniqueId(validity, 1), new UniqueId(validity, upper)));
            if (unreadOnly) query = query.And(SearchQuery.NotSeen);
            var found = await folder.SearchAsync(query, ct);
            var page = found.OrderByDescending(u => u.Id).Take(top).ToList();
            if (page.Count == 0) return new ListPage([], false);

            var items = MessageSummaryItems.UniqueId | MessageSummaryItems.Flags;
            if (Gmail) items |= MessageSummaryItems.GMailLabels;
            var summaries = await folder.FetchAsync(page, new FetchRequest(items), ct);
            var byUid = summaries.ToDictionary(s => s.UniqueId.Id);
            var refs = new List<MessageRef>(page.Count);
            foreach (var u in page)
            {
                if (!byUid.TryGetValue(u.Id, out var summary)) continue; // expunged between SEARCH and FETCH
                var flags = summary.Flags ?? MessageFlags.None;
                refs.Add(new MessageRef(
                    ImapSupport.FormatId(scope, validity, u.Id),
                    ImapSupport.FormatSortKey(validity, u.Id),
                    LabelsOf(summary),
                    Excluded: flags.HasFlag(MessageFlags.Draft) || flags.HasFlag(MessageFlags.Deleted),
                    IsRead: flags.HasFlag(MessageFlags.Seen)));
            }
            return new ListPage(refs, page.Count >= top);
        }, ct);

    private static void CheckValidity(IMailFolder folder, uint validity)
    {
        if (folder.UidValidity != validity)
            throw new MailboxException(false, "The mailbox was rebuilt (UIDVALIDITY changed); clear the session and start again.");
    }

    private IReadOnlyList<string> LabelsOf(IMessageSummary summary) => Gmail
        ? (summary.GMailLabels ?? []).ToList()
        : summary.Keywords is null ? [] : summary.Keywords.ToList();

    public async Task<List<ReadPart>> ReadMessagesAsync(IReadOnlyList<string> ids, ReadMode mode, CancellationToken ct)
    {
        var results = new ReadPart[ids.Count];
        var groups = new Dictionary<(string Scope, uint Validity), List<(int Index, uint Uid)>>();
        for (var i = 0; i < ids.Count; i++)
        {
            if (!ImapSupport.TryParseId(ids[i], out var scope, out var validity, out var uid))
            {
                results[i] = ReadPart.Fatal($"'{ids[i]}' is not an IMAP message id.");
                continue;
            }
            if (!groups.TryGetValue((scope, validity), out var list)) groups[(scope, validity)] = list = [];
            list.Add((i, uid));
        }

        foreach (var ((scope, validity), entries) in groups)
        {
            try
            {
                var parts = await WithClientAsync(client => ReadGroupAsync(client, scope, validity, entries.Select(e => e.Uid).ToList(), mode, ct), ct);
                for (var i = 0; i < entries.Count; i++) results[entries[i].Index] = parts[i];
            }
            catch (MailboxException ex)
            {
                var part = ex.IsTransient ? ReadPart.Transient(ex.Message)
                    : ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase) ? ReadPart.Unauthorized(ex.Message)
                    : ReadPart.Fatal(ex.Message);
                foreach (var (index, _) in entries) results[index] = part;
            }
        }
        return results.ToList();
    }

    private async Task<List<ReadPart>> ReadGroupAsync(ImapClient client, string scope, uint validity, List<uint> uids, ReadMode mode, CancellationToken ct)
    {
        var folder = await FolderAsync(client, scope, FolderAccess.ReadWrite, ct);
        CheckValidity(folder, validity);
        var wanted = uids.Select(u => new UniqueId(validity, u)).ToList();
        var output = new List<ReadPart>(uids.Count);

        switch (mode)
        {
            case ReadMode.Metadata:
            {
                var items = MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope | MessageSummaryItems.Flags |
                            MessageSummaryItems.BodyStructure | MessageSummaryItems.PreviewText;
                if (Gmail) items |= MessageSummaryItems.GMailLabels;
                var summaries = await folder.FetchAsync(wanted, new FetchRequest(items, ImapSupport.MetadataHeaders), ct);
                var byUid = summaries.ToDictionary(s => s.UniqueId.Id);
                foreach (var uid in uids)
                {
                    if (!byUid.TryGetValue(uid, out var s)) { output.Add(ReadPart.Missing()); continue; }
                    var labels = LabelsOf(s);
                    var metadata = ImapSupport.BuildMetadata(ImapSupport.FormatId(scope, validity, uid), s.Envelope, s.Headers,
                        s.Attachments.Any(), s.PreviewText, labels);
                    output.Add(new ReadPart { Ok = true, Metadata = metadata, Labels = labels });
                }
                break;
            }
            case ReadMode.Full:
            {
                foreach (var uid in wanted)
                {
                    MimeKit.MimeMessage? message;
                    try
                    {
                        message = await folder.GetMessageAsync(uid, ct);
                    }
                    catch (MessageNotFoundException)
                    {
                        output.Add(ReadPart.Missing());
                        continue;
                    }
                    var (text, reason) = ImapSupport.FullText(message);
                    output.Add(new ReadPart { Ok = true, FullText = text, FullReason = reason });
                }
                break;
            }
            default:
            {
                var items = MessageSummaryItems.UniqueId | MessageSummaryItems.Flags;
                if (Gmail) items |= MessageSummaryItems.GMailLabels;
                var summaries = await folder.FetchAsync(wanted, new FetchRequest(items), ct);
                var byUid = summaries.ToDictionary(s => s.UniqueId.Id);
                foreach (var uid in uids)
                {
                    output.Add(byUid.TryGetValue(uid, out var s) ? new ReadPart { Ok = true, Labels = LabelsOf(s) } : ReadPart.Missing());
                }
                break;
            }
        }
        return output;
    }

    /// <summary>Gmail: the label named after the rule. Generic IMAP: the rule's keyword (<see cref="ImapSupport.ToKeyword"/>).</summary>
    public string StoredLabel(string ruleName) => Gmail ? ruleName : ImapSupport.ToKeyword(ruleName);

    public Task<Dictionary<string, string>> EnsureLabelsAsync(IReadOnlyList<string> names, CancellationToken ct) =>
        WithClientAsync(async client =>
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (Gmail)
            {
                var root = client.GetFolder(client.PersonalNamespaces[0]);
                var existing = (await root.GetSubfoldersAsync(false, ct)).ToList();
                foreach (var name in names)
                {
                    var match = existing.FirstOrDefault(f => string.Equals(f.FullName, name, StringComparison.OrdinalIgnoreCase));
                    if (match is null)
                    {
                        match = await root.CreateAsync(name, true, ct) ?? throw new MailboxException(false, $"Gmail could not create the label '{name}'.");
                        existing.Add(match);
                    }
                    result[name.ToLowerInvariant()] = match.FullName;
                }
                return result;
            }

            await RequireKeywordsAsync(client, ct);
            foreach (var name in names) result[name.ToLowerInvariant()] = StoredLabel(name);
            return result;
        }, ct);

    private async Task RequireKeywordsAsync(ImapClient client, CancellationToken ct)
    {
        if (!_connectedOnce || !_keywordsSupported)
        {
            var inbox = await FolderAsync(client, ImapSupport.InboxScope, FolderAccess.ReadWrite, ct);
            _keywordsSupported = inbox.PermanentFlags.HasFlag(MessageFlags.UserDefined);
        }
        if (!_keywordsSupported)
            throw new MailboxException(false, "This IMAP server does not accept custom keywords; jevOutlook cannot label messages here.");
    }

    public async Task<WriteOutcome> ApplyLabelsAsync(IReadOnlyList<(string Id, IReadOnlyList<string> Labels)> updates, CancellationToken ct)
    {
        if (updates.Count == 0) return WriteOutcome.Success;
        try
        {
            await WithClientAsync<object?>(async client =>
            {
                if (!Gmail) await RequireKeywordsAsync(client, ct);
                foreach (var group in GroupIds(updates.Select(u => u.Id)))
                {
                    var folder = await FolderAsync(client, group.Key.Scope, FolderAccess.ReadWrite, ct);
                    CheckValidity(folder, group.Key.Validity);
                    var uids = group.Value.Select(g => new UniqueId(group.Key.Validity, g.Uid)).ToList();
                    var items = MessageSummaryItems.UniqueId | MessageSummaryItems.Flags;
                    if (Gmail) items |= MessageSummaryItems.GMailLabels;
                    var current = (await folder.FetchAsync(uids, new FetchRequest(items), ct)).ToDictionary(s => s.UniqueId.Id);

                    foreach (var (index, uid) in group.Value)
                    {
                        if (!current.TryGetValue(uid, out var summary)) continue; // vanished meanwhile: nothing to do
                        // Add-only (STORE +X-GM-LABELS / +FLAGS): never remove a label here, it may have been
                        // set by a filter or another client after the engine's re-read.
                        var add = ImapSupport.LabelsToAdd(LabelsOf(summary), updates[index].Labels);
                        if (add.Count == 0) continue;
                        var id = new UniqueId(group.Key.Validity, uid);
                        if (Gmail) await folder.StoreAsync(id, new StoreLabelsRequest(StoreAction.Add, add) { Silent = true }, ct);
                        else await folder.StoreAsync(id, new StoreFlagsRequest(StoreAction.Add, MessageFlags.None, add) { Silent = true }, ct);
                    }
                }
                return null;
            }, ct);
            return WriteOutcome.Success;
        }
        catch (MailboxException ex)
        {
            return new WriteOutcome(false, ex.IsTransient, ex.Message);
        }
    }

    public async Task<WriteOutcome> ArchiveAsync(IReadOnlyList<string> ids, CancellationToken ct)
    {
        if (ids.Count == 0) return WriteOutcome.Success;
        try
        {
            await WithClientAsync<object?>(async client =>
            {
                foreach (var group in GroupIds(ids))
                {
                    var folder = await FolderAsync(client, group.Key.Scope, FolderAccess.ReadWrite, ct);
                    CheckValidity(folder, group.Key.Validity);
                    var uids = group.Value.Select(g => new UniqueId(group.Key.Validity, g.Uid)).ToList();
                    // Only messages that still exist: MOVE on a vanished UID is a NO on some servers.
                    var present = (await folder.FetchAsync(uids, new FetchRequest(MessageSummaryItems.UniqueId), ct)).Select(s => s.UniqueId).ToList();
                    if (present.Count == 0) continue;

                    if (Gmail)
                    {
                        if (group.Key.Scope == ImapSupport.InboxScope)
                        {
                            var allMail = await AllMailAsync(client, ct);
                            await folder.MoveToAsync(present, allMail, ct);
                        }
                        else
                        {
                            await folder.StoreAsync(present, new StoreLabelsRequest(StoreAction.Remove, SystemLabels) { Silent = true }, ct);
                        }
                    }
                    else
                    {
                        var archive = await ArchiveFolderAsync(client, ct);
                        // Opening the archive folder (to create/find it) may have closed the Inbox: reopen.
                        var source = await FolderAsync(client, group.Key.Scope, FolderAccess.ReadWrite, ct);
                        await source.MoveToAsync(present, archive, ct);
                    }
                }
                return null;
            }, ct);
            return WriteOutcome.Success;
        }
        catch (MailboxException ex)
        {
            return new WriteOutcome(false, ex.IsTransient, ex.Message);
        }
    }

    private static Dictionary<(string Scope, uint Validity), List<(int Index, uint Uid)>> GroupIds(IEnumerable<string> ids)
    {
        var groups = new Dictionary<(string, uint), List<(int, uint)>>();
        var index = 0;
        foreach (var id in ids)
        {
            if (!ImapSupport.TryParseId(id, out var scope, out var validity, out var uid))
                throw new MailboxException(false, $"'{id}' is not an IMAP message id.");
            if (!groups.TryGetValue((scope, validity), out var list)) groups[(scope, validity)] = list = [];
            list.Add((index, uid));
            index++;
        }
        return groups;
    }

    public async Task<WriteOutcome> RemoveLabelAsync(IReadOnlyList<(string Id, IReadOnlyList<string> Labels)> messages, string label, CancellationToken ct)
    {
        if (messages.Count == 0) return WriteOutcome.Success;
        try
        {
            await WithClientAsync<object?>(async client =>
            {
                // The stored form (same mapping as FindMessagesWithLabelAsync) plus every spelling actually found
                // on the messages that matches it case-insensitively, so a case difference cannot leave it in place.
                var target = Gmail ? label : ImapSupport.ToKeyword(label);
                var spellings = messages.SelectMany(m => m.Labels)
                    .Where(l => string.Equals(l, target, StringComparison.OrdinalIgnoreCase))
                    .Append(target).Distinct(StringComparer.Ordinal).ToList();
                foreach (var group in GroupIds(messages.Select(m => m.Id)))
                {
                    var folder = await FolderAsync(client, group.Key.Scope, FolderAccess.ReadWrite, ct);
                    CheckValidity(folder, group.Key.Validity);
                    var uids = group.Value.Select(g => new UniqueId(group.Key.Validity, g.Uid)).ToList();
                    // Removing a label a message does not carry is a no-op; a vanished UID is ignored by STORE.
                    if (Gmail) await folder.StoreAsync(uids, new StoreLabelsRequest(StoreAction.Remove, spellings) { Silent = true }, ct);
                    else await folder.StoreAsync(uids, new StoreFlagsRequest(StoreAction.Remove, MessageFlags.None, spellings) { Silent = true }, ct);
                }
                return null;
            }, ct);
            return WriteOutcome.Success;
        }
        catch (MailboxException ex)
        {
            return new WriteOutcome(false, ex.IsTransient, ex.Message);
        }
    }

    public Task<List<(string Id, IReadOnlyList<string> Labels)>> FindMessagesWithLabelAsync(string label, int top, CancellationToken ct) =>
        WithClientAsync(async client =>
        {
            var scope = Gmail ? ImapSupport.AllScope : ImapSupport.InboxScope;
            var folder = await FolderAsync(client, scope, FolderAccess.ReadWrite, ct);
            var query = Gmail ? SearchQuery.HasGMailLabel(label) : SearchQuery.HasKeyword(ImapSupport.ToKeyword(label));
            var found = (await folder.SearchAsync(query, ct)).OrderByDescending(u => u.Id).Take(top).ToList();
            var result = new List<(string, IReadOnlyList<string>)>();
            if (found.Count == 0) return result;
            var items = MessageSummaryItems.UniqueId | MessageSummaryItems.Flags;
            if (Gmail) items |= MessageSummaryItems.GMailLabels;
            foreach (var s in await folder.FetchAsync(found, new FetchRequest(items), ct))
            {
                result.Add((ImapSupport.FormatId(scope, folder.UidValidity, s.UniqueId.Id), LabelsOf(s)));
            }
            return result;
        }, ct);

    public Task<bool> DeleteLabelAsync(string label, CancellationToken ct) =>
        WithClientAsync(async client =>
        {
            if (Gmail)
            {
                var folder = await FindLabelFolderAsync(client, label, ct);
                if (folder is null) return false;
                if (folder.IsOpen) await folder.CloseAsync(false, ct);
                await folder.DeleteAsync(ct);
                return true;
            }

            var keyword = ImapSupport.ToKeyword(label);
            var inbox = await FolderAsync(client, ImapSupport.InboxScope, FolderAccess.ReadWrite, ct);
            var any = false;
            var archive = await TryArchiveAsync(client, ct);
            foreach (var scopeFolder in archive is { } archiveFolder ? new[] { inbox, archiveFolder } : [inbox])
            {
                if (!scopeFolder.IsOpen || scopeFolder.Access < FolderAccess.ReadWrite) await scopeFolder.OpenAsync(FolderAccess.ReadWrite, ct);
                var found = await scopeFolder.SearchAsync(SearchQuery.HasKeyword(keyword), ct);
                if (found.Count == 0) continue;
                await scopeFolder.StoreAsync(found, new StoreFlagsRequest(StoreAction.Remove, MessageFlags.None, [keyword]) { Silent = true }, ct);
                any = true;
            }
            return any;
        }, ct);

    private static async Task<IMailFolder?> TryArchiveAsync(ImapClient client, CancellationToken ct)
    {
        if (client.Capabilities.HasFlag(ImapCapabilities.SpecialUse) || client.Capabilities.HasFlag(ImapCapabilities.XList))
        {
            try { if (client.GetFolder(SpecialFolder.Archive) is { } special) return special; }
            catch (Exception ex) when (ex is FolderNotFoundException or NotSupportedException) { }
        }
        foreach (var folder in await client.GetFoldersAsync(client.PersonalNamespaces[0], StatusItems.None, false, ct))
        {
            if (folder.Attributes.HasFlag(FolderAttributes.Archive) ||
                folder.Name.Equals("Archive", StringComparison.OrdinalIgnoreCase) || folder.Name.Equals("Archives", StringComparison.OrdinalIgnoreCase)) return folder;
        }
        return null;
    }

    // =====================================================================
    // Disposal
    // =====================================================================

    public void Dispose()
    {
        if (_client is { IsConnected: true })
        {
            try { _client.Disconnect(true); } catch (Exception) { /* best effort */ }
        }
        _client?.Dispose();
        _client = null;
        _gate.Dispose();
    }

    public async ValueTask DisposeAsync()
    {
        if (_client is { IsConnected: true })
        {
            try { await _client.DisconnectAsync(true); } catch (Exception) { /* best effort */ }
        }
        _client?.Dispose();
        _client = null;
        _gate.Dispose();
    }
}
