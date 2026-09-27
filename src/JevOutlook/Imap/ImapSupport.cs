using System.Globalization;
using System.Text;
using JevOutlook.Mail;
using MailKit;
using MimeKit;

namespace JevOutlook.Imap;

/// <summary>
/// Pure helpers behind <see cref="ImapMailbox"/>: cursor / id encoding, keyword
/// mapping, label diffs and MIME → <see cref="MessageMetadata"/> mapping. No I/O,
/// so everything here is unit-testable without a server.
/// </summary>
public static class ImapSupport
{
    public const string InboxScope = "inbox";
    public const string AllScope = "all";
    public const int SnippetChars = 255;

    // ----- Sort keys (cursor) and ids ------------------------------------------

    /// <summary>"validity:uid" with the UID zero-padded so lexicographic order equals numeric order.</summary>
    public static string FormatSortKey(uint validity, uint uid) =>
        string.Create(CultureInfo.InvariantCulture, $"{validity}:{uid:D12}");

    public static (uint Validity, uint Uid) ParseSortKey(string key)
    {
        var parts = (key ?? string.Empty).Split(':');
        if (parts.Length == 2 &&
            uint.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var validity) &&
            uint.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var uid))
        {
            return (validity, uid);
        }
        throw new MailboxException(false, $"The listing cursor '{key}' is not a valid IMAP position. Clear the session and start again.");
    }

    /// <summary>"scope:validity:uid" — self-describing so reads and writes know the folder.</summary>
    public static string FormatId(string scope, uint validity, uint uid) =>
        string.Create(CultureInfo.InvariantCulture, $"{scope}:{validity}:{uid}");

    public static bool TryParseId(string id, out string scope, out uint validity, out uint uid)
    {
        scope = string.Empty; validity = 0; uid = 0;
        var parts = (id ?? string.Empty).Split(':');
        if (parts.Length != 3 || parts[0] is not (InboxScope or AllScope)) return false;
        if (!uint.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out validity)) return false;
        if (!uint.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out uid) || uid == 0) return false;
        scope = parts[0];
        return true;
    }

    // ----- Keywords (generic IMAP) ------------------------------------------------

    /// <summary>
    /// IMAP keywords are atoms: no spaces, parentheses, braces, quotes, backslashes,
    /// '%', '*' or control characters. Rule names that already are atoms are used
    /// verbatim (so the engine's name comparison keeps matching); anything else is
    /// mapped character by character to '_'.
    /// </summary>
    public static string ToKeyword(string name)
    {
        var trimmed = (name ?? string.Empty).Trim();
        if (trimmed.Length == 0) return "_";
        var sb = new StringBuilder(trimmed.Length);
        foreach (var c in trimmed)
        {
            sb.Append(IsAtomChar(c) ? c : '_');
        }
        return sb.ToString();
    }

    public static bool IsAtom(string name) => name.Length > 0 && name.All(IsAtomChar);

    private static bool IsAtomChar(char c) =>
        c > ' ' && c < (char)127 && c is not ('(' or ')' or '{' or '"' or '\\' or '%' or '*' or ']' or '[');

    // ----- Label diffs ------------------------------------------------------------

    /// <summary>
    /// What to add and what to remove so that the message carries exactly
    /// <paramref name="desired"/> among user labels. System labels (starting
    /// with '\', e.g. Gmail's \Inbox, \Important) are never touched.
    /// </summary>
    public static (List<string> Add, List<string> Remove) Diff(IEnumerable<string> current, IEnumerable<string> desired)
    {
        var have = current.Where(l => !string.IsNullOrEmpty(l) && !l.StartsWith('\\')).ToList();
        var want = desired.Where(l => !string.IsNullOrEmpty(l) && !l.StartsWith('\\')).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var add = want.Where(w => !have.Any(h => string.Equals(h, w, StringComparison.OrdinalIgnoreCase))).ToList();
        var remove = have.Where(h => !want.Any(w => string.Equals(h, w, StringComparison.OrdinalIgnoreCase))).ToList();
        return (add, remove);
    }

    // ----- MIME → metadata --------------------------------------------------------

    /// <summary>Header fields fetched for the metadata stage (same set the Graph projection exposes).</summary>
    public static readonly string[] MetadataHeaders =
    [
        "List-Id", "List-Unsubscribe", "List-Unsubscribe-Post", "Auto-Submitted", "Precedence",
        "In-Reply-To", "References", "Importance", "X-Priority",
    ];

    public static MessageMetadata BuildMetadata(string id, Envelope? envelope, HeaderList? headers, bool hasAttachments,
        string? previewText, IReadOnlyList<string> labels, bool isDraft = false)
    {
        var h = headers ?? new HeaderList();
        var importance = Value(h, "Importance");
        if (importance.Length == 0) importance = PriorityToImportance(Value(h, "X-Priority"));
        return new MessageMetadata
        {
            Id = id,
            ConversationId = string.Empty,
            From = FormatAddresses(envelope?.From),
            To = FormatAddresses(envelope?.To),
            Cc = FormatAddresses(envelope?.Cc),
            ReplyTo = FormatAddresses(envelope?.ReplyTo),
            Subject = envelope?.Subject ?? string.Empty,
            Date = envelope?.Date is { } date ? date.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture) : string.Empty,
            Snippet = Snippet(previewText),
            ListId = Value(h, "List-Id"),
            ListUnsubscribe = Value(h, "List-Unsubscribe"),
            ListUnsubscribePost = Value(h, "List-Unsubscribe-Post"),
            AutoSubmitted = Value(h, "Auto-Submitted"),
            Precedence = Value(h, "Precedence"),
            InReplyTo = envelope?.InReplyTo ?? Value(h, "In-Reply-To"),
            References = Value(h, "References"),
            HasAttachments = hasAttachments,
            Importance = importance,
            InferenceClassification = string.Empty,
            Categories = labels,
        };
    }

    private static string Value(HeaderList headers, string field) => (headers[field] ?? string.Empty).Trim();

    private static string PriorityToImportance(string xPriority)
    {
        if (xPriority.Length == 0) return string.Empty;
        var digit = xPriority[0];
        return digit switch { '1' or '2' => "high", '4' or '5' => "low", '3' => "normal", _ => string.Empty };
    }

    /// <summary>"Name &lt;address&gt;" per mailbox, comma-joined — mirrors the Graph formatting.</summary>
    public static string FormatAddresses(InternetAddressList? list)
    {
        if (list is null || list.Count == 0) return string.Empty;
        var parts = new List<string>();
        foreach (var address in list.Mailboxes)
        {
            var name = (address.Name ?? string.Empty).Trim();
            var addr = (address.Address ?? string.Empty).Trim();
            if (name.Length == 0) { if (addr.Length > 0) parts.Add(addr); continue; }
            if (addr.Length == 0 || string.Equals(name, addr, StringComparison.OrdinalIgnoreCase)) { parts.Add(addr.Length > 0 ? addr : name); continue; }
            parts.Add($"{name} <{addr}>");
        }
        return string.Join(", ", parts);
    }

    /// <summary>The equivalent of Graph's bodyPreview: whitespace-normalised, first 255 characters.</summary>
    public static string Snippet(string? previewText)
    {
        if (string.IsNullOrWhiteSpace(previewText)) return string.Empty;
        var text = MessageContent.NormalizeWhitespace(previewText).Replace('\n', ' ');
        text = MessageContent.NormalizeWhitespace(text);
        return text.Length > SnippetChars ? text[..SnippetChars] : text;
    }

    /// <summary>Plain text of a full message: the text part first, else the HTML part converted.</summary>
    public static (string Text, string Reason) FullText(MimeMessage? message)
    {
        if (message is null) return (string.Empty, "The message has no body part.");
        var text = message.TextBody;
        if (string.IsNullOrWhiteSpace(text))
        {
            var html = message.HtmlBody;
            if (string.IsNullOrWhiteSpace(html)) return (string.Empty, "The message body is empty.");
            text = MessageContent.HtmlToText(html);
        }
        else if (MessageContent.LooksLikeHtml(text))
        {
            text = MessageContent.HtmlToText(text);
        }
        text = MessageContent.NormalizeWhitespace(text);
        return text.Length == 0 ? (string.Empty, "The message body contained no readable text.") : (text, string.Empty);
    }
}
