using System.Globalization;
using System.Text;
using MailClassification.Mail;
using MailKit;
using MimeKit;

namespace MailClassification.Imap;

/// <summary>
/// Pure helpers behind <see cref="ImapMailbox"/>: cursor / id encoding, keyword
/// mapping, add-only label writes and MIME → <see cref="MessageMetadata"/> mapping. No I/O,
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

    /// <summary>Prefix of the namespaced form given to reserved keyword names.</summary>
    public const string ReservedKeywordPrefix = "jev-";

    /// <summary>
    /// Keyword names that mail clients and servers give a meaning to. Thunderbird, Apple Mail
    /// and server-side antispam learners read <c>Junk</c> / <c>$Junk</c> / <c>NonJunk</c> /
    /// <c>$NotJunk</c> as the junk verdict (keywords compare case-insensitively), and
    /// <c>$Forwarded</c>, <c>$MDNSent</c>, <c>$Phishing</c>… as message state. MailClassification never
    /// marks mail as junk or changes such state, so a rule with one of these names is stored under
    /// <see cref="ReservedKeywordPrefix"/> instead. Every name starting with '$' (the IANA keyword
    /// registry's system-like range: $Label1..5, $Important, $Submitted…) is reserved as well.
    /// </summary>
    private static readonly HashSet<string> ReservedKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "junk", "nonjunk", "notjunk", "forwarded", "phishing", "mdnsent",
    };

    /// <summary>
    /// IMAP keywords are atoms: no spaces, parentheses, braces, quotes, backslashes,
    /// '%', '*' or control characters. Ordinary rule names that already are atoms are used
    /// verbatim (readable in mail clients); other characters are mapped to '_'. Reserved
    /// names (see <see cref="ReservedKeywords"/>) become <c>jev-&lt;name&gt;</c>, without the
    /// leading '$'. The mapping is idempotent: a keyword it produced maps to itself.
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
        var atom = sb.ToString();
        if (!atom.StartsWith('$') && !ReservedKeywords.Contains(atom)) return atom;
        var core = atom.TrimStart('$');
        return ReservedKeywordPrefix + (core.Length == 0 ? "_" : core);
    }

    public static bool IsAtom(string name) => name.Length > 0 && name.All(IsAtomChar);

    private static bool IsAtomChar(char c) =>
        c > ' ' && c < (char)127 && c is not ('(' or ')' or '{' or '"' or '\\' or '%' or '*' or ']' or '[');

    // ----- Label writes ------------------------------------------------------------

    /// <summary>
    /// Labels of <paramref name="desired"/> the message does not carry yet (case-insensitive).
    /// Label writes are add-only: a label present on the message but absent from
    /// <paramref name="desired"/> may have been added a moment ago by a Gmail filter or another
    /// client, so it is never removed here (removal is the explicit cleanup path). Labels are in
    /// stored form (Gmail label / IMAP keyword) and are not re-mapped: a keyword a client set,
    /// such as "$Forwarded", must not turn into a new one. System labels (starting with '\',
    /// e.g. Gmail's \Inbox, \Spam) are never added.
    /// </summary>
    public static List<string> LabelsToAdd(IEnumerable<string> current, IEnumerable<string> desired)
    {
        var have = current.Where(l => !string.IsNullOrEmpty(l)).ToList();
        return desired
            .Where(l => !string.IsNullOrEmpty(l) && !l.StartsWith('\\'))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(w => !have.Any(h => string.Equals(h, w, StringComparison.OrdinalIgnoreCase)))
            .ToList();
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
