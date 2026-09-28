using System.Text.Json;
using MailClassification;
using MailClassification.Graph;
using MailClassification.Mail;

namespace MailClassification.Tests;

public class MessageContentTests
{
    [Fact]
    public void Compact_keeps_short_bodies_untouched()
    {
        Assert.Equal("hello", MessageContent.Compact("  hello \n"));
    }

    [Fact]
    public void Compact_keeps_head_and_tail_of_long_bodies()
    {
        var text = new string('a', 4000) + new string('z', 4000);
        var compact = MessageContent.Compact(text);
        Assert.True(compact.Length <= AppConstants.MaxBodyChars);
        Assert.StartsWith("aaaa", compact);
        Assert.EndsWith("zzzz", compact);
        Assert.Contains("[Middle of long message omitted]", compact);
    }

    [Fact]
    public void Html_is_converted_to_readable_text()
    {
        const string html = "<html><head><style>p{color:red}</style><script>alert(1)</script></head><body><p>Hello&nbsp;<b>World</b></p><div>Second &amp; line</div><!-- c --></body></html>";
        var text = MessageContent.NormalizeWhitespace(MessageContent.HtmlToText(html));
        Assert.DoesNotContain("alert", text);
        Assert.DoesNotContain("color", text);
        Assert.Contains("Hello World", text);
        Assert.Contains("Second & line", text);
    }

    [Fact]
    public void Extract_reports_missing_and_empty_bodies()
    {
        using var noBody = JsonDocument.Parse("{\"id\":\"1\"}");
        Assert.Equal(string.Empty, MessageContent.Extract(noBody.RootElement).Text);
        using var empty = JsonDocument.Parse("{\"body\":{\"contentType\":\"text\",\"content\":\"   \"}}");
        Assert.NotEqual(string.Empty, MessageContent.Extract(empty.RootElement).Reason);
        using var tagsOnly = JsonDocument.Parse("{\"body\":{\"contentType\":\"html\",\"content\":\"<div><img src='x'></div>\"}}");
        var (text, reason) = MessageContent.Extract(tagsOnly.RootElement);
        Assert.Equal(string.Empty, text);
        Assert.Contains("no readable text", reason);
    }

    [Fact]
    public void Metadata_is_normalized_from_graph_message()
    {
        const string json = """
        {
          "id": "AAMk",
          "conversationId": "conv",
          "subject": "Invoice",
          "from": {"emailAddress": {"name": "Billing", "address": "billing@example.com"}},
          "toRecipients": [{"emailAddress": {"name": "", "address": "me@example.com"}}, {"emailAddress": {"name": "Bob", "address": "bob@example.com"}}],
          "ccRecipients": [],
          "replyTo": [{"emailAddress": {"name": "noreply@example.com", "address": "noreply@example.com"}}],
          "receivedDateTime": "2026-09-22T08:00:00Z",
          "bodyPreview": "Your invoice is attached",
          "categories": ["Blue category", "jev-triaged"],
          "hasAttachments": true,
          "importance": "normal",
          "inferenceClassification": "other",
          "internetMessageHeaders": [
            {"name": "List-Unsubscribe", "value": "<mailto:u@example.com>"},
            {"name": "Auto-Submitted", "value": "auto-generated"},
            {"name": "Auto-Submitted", "value": "duplicate ignored"}
          ]
        }
        """;
        using var doc = JsonDocument.Parse(json);
        var m = MessageMetadata.FromGraph(doc.RootElement);
        Assert.Equal("AAMk", m.Id);
        Assert.Equal("Billing <billing@example.com>", m.From);
        Assert.Equal("me@example.com, Bob <bob@example.com>", m.To);
        Assert.Equal(string.Empty, m.Cc);
        Assert.Equal("noreply@example.com", m.ReplyTo);
        Assert.Equal("Your invoice is attached", m.Snippet);
        Assert.Equal("<mailto:u@example.com>", m.ListUnsubscribe);
        Assert.Equal("auto-generated", m.AutoSubmitted);
        Assert.True(m.HasAttachments);
        Assert.Equal("other", m.InferenceClassification);
        Assert.Equal(["Blue category", "jev-triaged"], m.Categories);
    }
}
