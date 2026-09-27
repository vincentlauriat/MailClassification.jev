using JevOutlook.Imap;
using JevOutlook.Mail;
using JevOutlook.Storage;
using MailKit;
using MimeKit;

namespace JevOutlook.Tests;

public class ImapMailboxTests
{
    // ----- sort keys / ids -------------------------------------------------

    [Fact]
    public void Sort_key_orders_lexicographically_like_uids()
    {
        var a = ImapSupport.FormatSortKey(7, 12);
        var b = ImapSupport.FormatSortKey(7, 100);
        var c = ImapSupport.FormatSortKey(7, 4_000_000_000);
        Assert.Equal("7:000000000012", a);
        Assert.True(string.CompareOrdinal(a, b) < 0);
        Assert.True(string.CompareOrdinal(b, c) < 0);
        Assert.Equal((7u, 100u), ImapSupport.ParseSortKey(b));
        Assert.Equal((7u, 4_000_000_000u), ImapSupport.ParseSortKey(c));
    }

    [Fact]
    public void Invalid_sort_key_is_a_non_transient_mailbox_error()
    {
        var ex = Assert.Throws<MailboxException>(() => ImapSupport.ParseSortKey("2026-09-22T10:30:15.0000000Z"));
        Assert.False(ex.IsTransient);
        Assert.Throws<MailboxException>(() => ImapSupport.ParseSortKey(string.Empty));
    }

    [Fact]
    public void Ids_round_trip_and_reject_garbage()
    {
        var id = ImapSupport.FormatId("inbox", 42, 1234);
        Assert.Equal("inbox:42:1234", id);
        Assert.True(ImapSupport.TryParseId(id, out var scope, out var validity, out var uid));
        Assert.Equal(("inbox", 42u, 1234u), (scope, validity, uid));

        Assert.False(ImapSupport.TryParseId("AAMkAGI2...", out _, out _, out _));      // a Graph id
        Assert.False(ImapSupport.TryParseId("junk:42:1", out _, out _, out _));         // unknown scope
        Assert.False(ImapSupport.TryParseId("inbox:42:0", out _, out _, out _));        // uid 0 is not valid
        Assert.False(ImapSupport.TryParseId("inbox:42", out _, out _, out _));
    }

    // ----- keywords -----------------------------------------------------------

    [Theory]
    [InlineData("action-required", "action-required")]
    [InlineData("newsletters", "newsletters")]
    [InlineData("Cold outreach", "Cold_outreach")]
    [InlineData("a(b)c*d\"e\\f%g", "a_b_c_d_e_f_g")]
    [InlineData("  spaced  ", "spaced")]
    [InlineData("", "_")]
    public void Rule_names_map_to_imap_atoms(string name, string expected)
    {
        var keyword = ImapSupport.ToKeyword(name);
        Assert.Equal(expected, keyword);
        Assert.True(ImapSupport.IsAtom(keyword));
    }

    // ----- label diff -------------------------------------------------------

    [Fact]
    public void Diff_adds_missing_removes_extra_and_ignores_system_labels()
    {
        var (add, remove) = ImapSupport.Diff(
            current: ["\\Inbox", "\\Important", "newsletters", "Old-Label"],
            desired: ["Newsletters", "action-required"]);
        Assert.Equal(["action-required"], add);        // "newsletters" already there (case-insensitive)
        Assert.Equal(["Old-Label"], remove);           // system labels untouched
    }

    [Fact]
    public void Diff_is_a_no_op_when_labels_already_match()
    {
        var (add, remove) = ImapSupport.Diff(["\\Inbox", "newsletters"], ["newsletters", "\\Inbox"]);
        Assert.Empty(add);
        Assert.Empty(remove);
    }

    // ----- metadata mapping ----------------------------------------------------

    [Fact]
    public void Envelope_and_headers_map_to_the_graph_shaped_metadata()
    {
        var envelope = new Envelope
        {
            Subject = "Weekly digest",
            Date = new DateTimeOffset(2026, 9, 24, 10, 30, 0, TimeSpan.FromHours(2)),
            InReplyTo = "<parent@example.org>",
        };
        envelope.From.Add(new MailboxAddress("TLDR", "dan@tldrnewsletter.com"));
        envelope.To.Add(new MailboxAddress(string.Empty, "vincent@example.com"));
        envelope.To.Add(new MailboxAddress("Vincent", "vincent@example.com"));
        envelope.Cc.Add(new MailboxAddress("same@example.com", "same@example.com"));
        var headers = new HeaderList
        {
            { "List-Id", "<tldr.list-id.net>" },
            { "List-Unsubscribe", "<https://example.org/u>" },
            { "Precedence", "bulk" },
            { "X-Priority", "1 (Highest)" },
        };

        var metadata = ImapSupport.BuildMetadata("inbox:1:5", envelope, headers, hasAttachments: true,
            previewText: "  Hello\n\n\n   world  ", labels: ["newsletters"]);

        Assert.Equal("inbox:1:5", metadata.Id);
        Assert.Equal("TLDR <dan@tldrnewsletter.com>", metadata.From);
        Assert.Equal("vincent@example.com, Vincent <vincent@example.com>", metadata.To);
        Assert.Equal("same@example.com", metadata.Cc);
        Assert.Equal("Weekly digest", metadata.Subject);
        Assert.Equal("2026-09-24T08:30:00Z", metadata.Date);      // UTC
        Assert.Equal("Hello world", metadata.Snippet);
        Assert.Equal("<tldr.list-id.net>", metadata.ListId);
        Assert.Equal("<https://example.org/u>", metadata.ListUnsubscribe);
        Assert.Equal("bulk", metadata.Precedence);
        Assert.Equal("<parent@example.org>", metadata.InReplyTo);
        Assert.Equal("high", metadata.Importance);                // from X-Priority
        Assert.True(metadata.HasAttachments);
        Assert.Equal(["newsletters"], metadata.Categories);
    }

    [Fact]
    public void Missing_envelope_yields_empty_but_valid_metadata()
    {
        var metadata = ImapSupport.BuildMetadata("inbox:1:9", null, null, false, null, []);
        Assert.Equal("inbox:1:9", metadata.Id);
        Assert.Equal(string.Empty, metadata.From);
        Assert.Equal(string.Empty, metadata.Date);
        Assert.Equal(string.Empty, metadata.Snippet);
        Assert.Empty(metadata.Categories);
    }

    [Fact]
    public void Snippet_is_normalised_and_cut_like_body_preview()
    {
        var text = string.Join(" ", Enumerable.Repeat("word", 100));
        var snippet = ImapSupport.Snippet(text);
        Assert.Equal(ImapSupport.SnippetChars, snippet.Length);
        Assert.Equal(string.Empty, ImapSupport.Snippet("   \n\t "));
        Assert.Equal("a b", ImapSupport.Snippet("a\r\n\r\nb"));
    }

    // ----- full text ------------------------------------------------------------

    [Fact]
    public void Full_text_prefers_plain_then_converts_html()
    {
        var plain = new MimeMessage { Body = new TextPart("plain") { Text = "Hello  there\n\n\n\nBye" } };
        Assert.Equal(("Hello there\n\nBye", string.Empty), ImapSupport.FullText(plain));

        var html = new MimeMessage { Body = new TextPart("html") { Text = "<html><body><p>Hi&nbsp;<b>you</b></p><script>x()</script></body></html>" } };
        var (text, reason) = ImapSupport.FullText(html);
        Assert.Equal(string.Empty, reason);
        Assert.Equal("Hi you", text);

        var empty = new MimeMessage { Body = new TextPart("plain") { Text = "   " } };
        Assert.NotEqual(string.Empty, ImapSupport.FullText(empty).Reason);
        Assert.NotEqual(string.Empty, ImapSupport.FullText(null).Reason);
    }

    // ----- capabilities ---------------------------------------------------------

    [Fact]
    public void Capabilities_follow_the_account_kind()
    {
        var gmail = new ImapMailbox(new MailAccount { Id = "g", Email = "x@gmail.com", Kind = AccountKind.Imap, Gmail = true, Host = "imap.gmail.com" }, "pw");
        Assert.Equal("Gmail", gmail.Capabilities.ProviderName);
        Assert.Equal("labels", gmail.Capabilities.LabelNounPlural);
        Assert.True(gmail.Capabilities.SupportsAllScope);

        var generic = new ImapMailbox(new MailAccount { Id = "i", Email = "x@example.org", Kind = AccountKind.Imap, Host = "mail.example.org" }, "pw");
        Assert.Equal("IMAP", generic.Capabilities.ProviderName);
        Assert.Equal("keyword", generic.Capabilities.LabelNoun);
        Assert.False(generic.Capabilities.SupportsAllScope);
        Assert.True(generic.Capabilities.CanArchive);
        gmail.Dispose();
        generic.Dispose();
    }

    [Fact]
    public async Task Reading_a_non_imap_id_fails_closed_without_connecting()
    {
        using var mailbox = new ImapMailbox(new MailAccount { Id = "i", Email = "x@example.org", Kind = AccountKind.Imap, Host = "mail.example.org" }, "pw");
        var parts = await mailbox.ReadMessagesAsync(["AAMkAGI2-not-imap"], ReadMode.Metadata, CancellationToken.None);
        Assert.Single(parts);
        Assert.False(parts[0].Ok);
        Assert.False(parts[0].Retryable);
        Assert.False(parts[0].NotFound);
    }
}
