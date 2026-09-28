using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JevOutlook.Mail;

/// <summary>
/// Normalized, cheap view of one Outlook message: the equivalent of the Gmail
/// "metadata" format (headers + snippet). Built from a Graph message resource
/// selected with the Graph metadata projection (or from IMAP headers).
/// </summary>
public sealed class MessageMetadata
{
    public string Id { get; init; } = string.Empty;
    public string ConversationId { get; init; } = string.Empty;
    public string From { get; init; } = string.Empty;
    public string To { get; init; } = string.Empty;
    public string Cc { get; init; } = string.Empty;
    public string ReplyTo { get; init; } = string.Empty;
    public string Subject { get; init; } = string.Empty;
    public string Date { get; init; } = string.Empty;
    public string Snippet { get; init; } = string.Empty;
    public string ListId { get; init; } = string.Empty;
    public string ListUnsubscribe { get; init; } = string.Empty;
    public string ListUnsubscribePost { get; init; } = string.Empty;
    public string AutoSubmitted { get; init; } = string.Empty;
    public string Precedence { get; init; } = string.Empty;
    public string InReplyTo { get; init; } = string.Empty;
    public string References { get; init; } = string.Empty;
    public bool HasAttachments { get; init; }
    public string Importance { get; init; } = string.Empty;
    public string InferenceClassification { get; init; } = string.Empty;
    public IReadOnlyList<string> Categories { get; init; } = [];

    public static MessageMetadata Empty(string id) => new() { Id = id };

    public static MessageMetadata FromGraph(JsonElement message)
    {
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (message.TryGetProperty("internetMessageHeaders", out var headerList) && headerList.ValueKind == JsonValueKind.Array)
        {
            foreach (var header in headerList.EnumerateArray())
            {
                var name = GetString(header, "name");
                if (name.Length == 0 || headers.ContainsKey(name)) continue; // keep first occurrence
                headers[name] = GetString(header, "value");
            }
        }

        return new MessageMetadata
        {
            Id = GetString(message, "id"),
            ConversationId = GetString(message, "conversationId"),
            From = FormatRecipient(message, "from"),
            To = FormatRecipients(message, "toRecipients"),
            Cc = FormatRecipients(message, "ccRecipients"),
            ReplyTo = FormatRecipients(message, "replyTo"),
            Subject = GetString(message, "subject"),
            Date = GetString(message, "receivedDateTime"),
            Snippet = GetString(message, "bodyPreview"),
            ListId = headers.GetValueOrDefault("List-Id", string.Empty),
            ListUnsubscribe = headers.GetValueOrDefault("List-Unsubscribe", string.Empty),
            ListUnsubscribePost = headers.GetValueOrDefault("List-Unsubscribe-Post", string.Empty),
            AutoSubmitted = headers.GetValueOrDefault("Auto-Submitted", string.Empty),
            Precedence = headers.GetValueOrDefault("Precedence", string.Empty),
            InReplyTo = headers.GetValueOrDefault("In-Reply-To", string.Empty),
            References = headers.GetValueOrDefault("References", string.Empty),
            HasAttachments = message.TryGetProperty("hasAttachments", out var att) && att.ValueKind == JsonValueKind.True,
            Importance = GetString(message, "importance"),
            InferenceClassification = GetString(message, "inferenceClassification"),
            Categories = ReadCategories(message),
        };
    }

    public static IReadOnlyList<string> ReadCategories(JsonElement message)
    {
        if (!message.TryGetProperty("categories", out var categories) || categories.ValueKind != JsonValueKind.Array) return [];
        return categories.EnumerateArray()
            .Where(c => c.ValueKind == JsonValueKind.String)
            .Select(c => c.GetString() ?? string.Empty)
            .Where(c => c.Length > 0)
            .ToList();
    }

    internal static string GetString(JsonElement element, string property)
    {
        if (element.ValueKind != JsonValueKind.Object) return string.Empty;
        if (!element.TryGetProperty(property, out var value)) return string.Empty;
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? string.Empty,
            JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => value.GetRawText(),
            _ => string.Empty,
        };
    }

    private static string FormatRecipient(JsonElement message, string property)
    {
        if (!message.TryGetProperty(property, out var recipient)) return string.Empty;
        return FormatEmailAddress(recipient);
    }

    private static string FormatRecipients(JsonElement message, string property)
    {
        if (!message.TryGetProperty(property, out var list) || list.ValueKind != JsonValueKind.Array) return string.Empty;
        return string.Join(", ", list.EnumerateArray().Select(FormatEmailAddress).Where(s => s.Length > 0));
    }

    private static string FormatEmailAddress(JsonElement recipient)
    {
        if (recipient.ValueKind != JsonValueKind.Object) return string.Empty;
        if (!recipient.TryGetProperty("emailAddress", out var email) || email.ValueKind != JsonValueKind.Object) return string.Empty;
        var name = GetString(email, "name").Trim();
        var address = GetString(email, "address").Trim();
        if (name.Length == 0) return address;
        if (address.Length == 0 || string.Equals(name, address, StringComparison.OrdinalIgnoreCase)) return address.Length > 0 ? address : name;
        return $"{name} <{address}>";
    }
}

/// <summary>Plain-text extraction from a full Graph message (body requested as text).</summary>
public static class MessageContent
{
    private static readonly Regex ScriptStyle = new("<(script|style)[^>]*>.*?</\\1>", RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex Comments = new("<!--.*?-->", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex BlockBreaks = new("<\\s*(br|/p|/div|/li|/tr|/h[1-6]|/blockquote|/pre)\\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex Tags = new("<[^>]+>", RegexOptions.Compiled);
    private static readonly Regex Spaces = new("[ \\t\\u00a0]+", RegexOptions.Compiled);
    private static readonly Regex Newlines = new("\\n{3,}", RegexOptions.Compiled);

    public static (string Text, string Reason) Extract(JsonElement full)
    {
        if (!full.TryGetProperty("body", out var body) || body.ValueKind != JsonValueKind.Object)
        {
            return (string.Empty, "The message has no body part.");
        }
        var contentType = MessageMetadata.GetString(body, "contentType");
        var content = MessageMetadata.GetString(body, "content");
        if (string.IsNullOrWhiteSpace(content)) return (string.Empty, "The message body is empty.");

        var text = string.Equals(contentType, "html", StringComparison.OrdinalIgnoreCase) || LooksLikeHtml(content)
            ? HtmlToText(content)
            : content;
        text = NormalizeWhitespace(text);
        return text.Length == 0
            ? (string.Empty, "The message body contained no readable text.")
            : (text, string.Empty);
    }

    public static bool LooksLikeHtml(string content) =>
        content.Contains("<html", StringComparison.OrdinalIgnoreCase) ||
        content.Contains("<body", StringComparison.OrdinalIgnoreCase) ||
        content.Contains("<div", StringComparison.OrdinalIgnoreCase);

    public static string HtmlToText(string html)
    {
        var text = ScriptStyle.Replace(html, " ");
        text = Comments.Replace(text, " ");
        text = BlockBreaks.Replace(text, "\n");
        text = Tags.Replace(text, " ");
        text = WebUtility.HtmlDecode(text);
        return text;
    }

    public static string NormalizeWhitespace(string text)
    {
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var sb = new StringBuilder(text.Length);
        foreach (var line in lines)
        {
            sb.Append(Spaces.Replace(line, " ").Trim()).Append('\n');
        }
        return Newlines.Replace(sb.ToString(), "\n\n").Trim();
    }

    /// <summary>Keep head and tail of very long bodies (same policy as jevMail).</summary>
    public static string Compact(string text)
    {
        var normalized = (text ?? string.Empty).Trim();
        if (normalized.Length <= AppConstants.MaxBodyChars) return normalized;
        var tailLength = Math.Min(800, (int)Math.Floor(AppConstants.MaxBodyChars * 0.2));
        const string marker = "\n\n[Middle of long message omitted]\n\n";
        var headLength = Math.Max(0, AppConstants.MaxBodyChars - tailLength - marker.Length);
        return normalized[..headLength].TrimEnd() + marker + normalized[^tailLength..].TrimStart();
    }
}
