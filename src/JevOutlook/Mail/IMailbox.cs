namespace JevOutlook.Mail;

/// <summary>One message as seen by the listing pass (cheap: no headers).</summary>
/// <param name="Id">Provider-specific stable id, opaque to the engine.</param>
/// <param name="SortKey">Newest-first ordering key (Graph: receivedDateTime ISO-8601; IMAP: zero-padded UID). Equal keys are ties.</param>
/// <param name="Labels">Labels / categories / keywords already on the message.</param>
/// <param name="Excluded">True for junk, deleted or draft items that a whole-mailbox scope must skip.</param>
public sealed record MessageRef(string Id, string SortKey, IReadOnlyList<string> Labels, bool Excluded, bool IsRead);

public sealed record ListPage(IReadOnlyList<MessageRef> Items, bool Full);

public enum ReadMode { Metadata, Full, Labels }

/// <summary>Result of reading one message; mirrors one inner response of a batch.</summary>
public sealed class ReadPart
{
    public bool Ok { get; init; }
    /// <summary>The message vanished (moved or deleted since it was listed): nothing to do, not an error.</summary>
    public bool NotFound { get; init; }
    public bool Retryable { get; init; }
    public bool AuthFailure { get; init; }
    public int RetryAfterMs { get; init; }
    public string Error { get; init; } = string.Empty;

    /// <summary><see cref="ReadMode.Metadata"/>: headers + preview.</summary>
    public MessageMetadata? Metadata { get; init; }
    /// <summary><see cref="ReadMode.Full"/>: plain text already extracted from the body (empty when unreadable).</summary>
    public string FullText { get; init; } = string.Empty;
    /// <summary><see cref="ReadMode.Full"/>: why <see cref="FullText"/> is empty.</summary>
    public string FullReason { get; init; } = string.Empty;
    /// <summary><see cref="ReadMode.Labels"/> (and Metadata): current labels.</summary>
    public IReadOnlyList<string> Labels { get; init; } = [];

    public static ReadPart Missing() => new() { Ok = false, NotFound = true, Error = "Message not found." };
    public static ReadPart Transient(string error, int retryAfterMs = 0) => new() { Ok = false, Retryable = true, RetryAfterMs = retryAfterMs, Error = error };
    public static ReadPart Fatal(string error) => new() { Ok = false, Retryable = false, Error = error };
    public static ReadPart Unauthorized(string error) => new() { Ok = false, AuthFailure = true, Error = error };
}

public sealed record WriteOutcome(bool Ok, bool Retryable, string Error)
{
    public static readonly WriteOutcome Success = new(true, false, string.Empty);
}

/// <summary>What the connected mailbox can do and how to talk about it to the user.</summary>
/// <param name="ProviderName">"Outlook", "Gmail" or "IMAP".</param>
/// <param name="LabelNoun">"category" / "label" / "keyword".</param>
/// <param name="CanArchive">False when the provider has no archive destination for this mailbox.</param>
/// <param name="SupportsAllScope">False when only the Inbox can be scanned (generic IMAP).</param>
/// <param name="ArchiveDescription">Human sentence describing what archiving does here.</param>
public sealed record MailboxCapabilities(string ProviderName, string LabelNoun, string LabelNounPlural, bool CanArchive, bool SupportsAllScope, string ArchiveDescription);

/// <summary>
/// The exact mailbox surface the triage engine needs, so Outlook (Graph), Gmail
/// (IMAP + X-GM-LABELS) and generic IMAP (keywords) are interchangeable.
/// Message ids and cursors are opaque strings owned by the provider.
/// </summary>
public interface IMailbox
{
    MailboxCapabilities Capabilities { get; }

    /// <summary>"Name &lt;address&gt;" of the mailbox owner.</summary>
    Task<string> GetIdentityAsync(CancellationToken ct);

    /// <summary>Resolve folders, probe capabilities. Idempotent; cheap after the first call.</summary>
    Task<MailboxCapabilities> ConnectAsync(CancellationToken ct);

    /// <summary>Approximate number of candidate messages, or -1 when unknown.</summary>
    Task<long> EstimateAsync(string scope, bool unreadOnly, CancellationToken ct);

    /// <summary>A sort key strictly newer than any existing message (the listing starts from it).</summary>
    Task<string> GetInitialCursorAsync(string scope, CancellationToken ct);

    /// <summary>Newest-first page with SortKey ≤ cursor (or &lt; when <paramref name="exclusive"/>).</summary>
    Task<ListPage> ListMessagesAsync(string scope, bool unreadOnly, string cursor, bool exclusive, int top, CancellationToken ct);

    /// <summary>Read many messages; result order matches <paramref name="ids"/>.</summary>
    Task<List<ReadPart>> ReadMessagesAsync(IReadOnlyList<string> ids, ReadMode mode, CancellationToken ct);

    /// <summary>Make sure every label exists (with a colour where the provider has one). Returns canonical names keyed by lower-case name.</summary>
    Task<Dictionary<string, string>> EnsureLabelsAsync(IReadOnlyList<string> names, CancellationToken ct);

    /// <summary>Replace the label set of each message (idempotent; a vanished message counts as done).</summary>
    Task<WriteOutcome> ApplyLabelsAsync(IReadOnlyList<(string Id, IReadOnlyList<string> Labels)> updates, CancellationToken ct);

    /// <summary>Archive messages (provider-specific: move to Archive, remove from Inbox…). Idempotent.</summary>
    Task<WriteOutcome> ArchiveAsync(IReadOnlyList<string> ids, CancellationToken ct);

    /// <summary>Ids (+ current labels) of up to <paramref name="top"/> messages carrying <paramref name="label"/>.</summary>
    Task<List<(string Id, IReadOnlyList<string> Labels)>> FindMessagesWithLabelAsync(string label, int top, CancellationToken ct);

    /// <summary>Delete a label definition (messages keep the name until rewritten). False when it did not exist.</summary>
    Task<bool> DeleteLabelAsync(string label, CancellationToken ct);
}

/// <summary>A mailbox-level failure. Transient ones pause the job; others fail it.</summary>
public class MailboxException(bool transient, string message) : Exception(message)
{
    public bool IsTransient { get; } = transient;
}
