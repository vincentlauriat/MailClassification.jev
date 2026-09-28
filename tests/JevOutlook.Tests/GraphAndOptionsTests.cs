using JevOutlook;
using JevOutlook.Graph;
using JevOutlook.Mail;
using JevOutlook.Storage;
using JevOutlook.Triage;

namespace JevOutlook.Tests;

public class GraphAndOptionsTests
{
    [Fact]
    public void Filter_puts_orderby_property_first_and_uses_utc()
    {
        var cursor = new DateTimeOffset(2026, 9, 22, 12, 30, 15, TimeSpan.FromHours(2));
        Assert.Equal("receivedDateTime le 2026-09-22T10:30:15.0000000Z", GraphMailClient.BuildFilter(false, cursor));
        Assert.Equal("receivedDateTime le 2026-09-22T10:30:15.0000000Z and isRead eq false", GraphMailClient.BuildFilter(true, cursor));
        Assert.Equal("receivedDateTime lt 2026-09-22T10:30:15.0000000Z", GraphMailClient.BuildFilter(false, cursor, exclusive: true));
    }

    [Fact]
    public void Mailboxes_are_async_disposable_so_one_shot_connections_are_closed()
    {
        Assert.True(typeof(IAsyncDisposable).IsAssignableFrom(typeof(IMailbox)));
    }

    private sealed class OkHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.OK));
    }

    [Fact]
    public async Task Disposing_a_graph_mailbox_keeps_the_shared_http_client()
    {
        using var http = new HttpClient(new OkHandler());
        var account = new MailAccount { Id = "m", Email = "m@contoso.com", Kind = AccountKind.Graph };
        var config = new AppConfig { ClientId = "00000000-0000-0000-0000-000000000001" };
        object mailbox = new GraphMailClient(http, new GraphTokenProvider(config, account, null, deviceCode: true));
        await ((IAsyncDisposable)mailbox).DisposeAsync();
        using var response = await http.GetAsync("http://127.0.0.1/still-usable");
        Assert.True(response.IsSuccessStatusCode);
    }

    [Fact]
    public void Retry_after_parses_seconds_and_dates()
    {
        Assert.Equal(2500, GraphMailClient.ParseRetryAfter("2.5"));
        Assert.Equal(0, GraphMailClient.ParseRetryAfter(""));
        Assert.Equal(0, GraphMailClient.ParseRetryAfter("garbage"));
        Assert.True(GraphMailClient.ParseRetryAfter(DateTimeOffset.UtcNow.AddSeconds(5).ToString("R")) > 1000);
    }

    [Theory]
    [InlineData(429, true)]
    [InlineData(503, true)]
    [InlineData(0, true)]
    [InlineData(404, false)]
    [InlineData(400, false)]
    public void Transient_statuses(int status, bool transient) => Assert.Equal(transient, GraphMailClient.IsTransient(status));

    [Fact]
    public void Well_known_folders_exclude_junk_deleted_and_drafts()
    {
        var folders = new WellKnownFolders("inbox", "archive", "junk", "deleted", "drafts", "sent");
        Assert.True(folders.IsExcluded("junk"));
        Assert.True(folders.IsExcluded("drafts"));
        Assert.False(folders.IsExcluded("sent"));
        Assert.False(folders.IsExcluded("inbox"));
    }

    [Fact]
    public void Run_options_defaults_mirror_the_gmail_ui()
    {
        var options = RunOptions.Normalize(null, true, true, null, null,
            AppConstants.DefaultMaxSpendUsd, AppConstants.DefaultMetadataThreshold, AppConstants.DefaultArchiveThreshold);
        Assert.Equal("inbox", options.Scope);
        Assert.Equal(RunMode.Labels, options.Mode);
        Assert.Equal(10, options.Limit);
        Assert.True(options.DryRun);
        Assert.True(options.UnreadOnly);
    }

    [Fact]
    public void Run_options_validate_ranges()
    {
        Assert.Throws<ArgumentException>(() => RunOptions.Normalize("inbox", true, true, "labels", "0", 0.1, 0.75, 0.93));
        Assert.Throws<ArgumentException>(() => RunOptions.Normalize("inbox", true, true, "labels", "ten", 0.1, 0.75, 0.93));
        Assert.Throws<ArgumentException>(() => RunOptions.Normalize("inbox", true, true, "labels", "10", 0, 0.75, 0.93));
        Assert.Throws<ArgumentException>(() => RunOptions.Normalize("inbox", true, true, "labels", "10", 0.1, 0.3, 0.93));
        Assert.Throws<ArgumentException>(() => RunOptions.Normalize("inbox", true, true, "labels", "10", 0.1, 0.9, 0.8));
        var all = RunOptions.Normalize("ALL", false, false, "labels-archive", "all", 5, 0.8, 0.95);
        Assert.Null(all.Limit);
        Assert.Equal("all", all.Scope);
        Assert.Equal(RunMode.LabelsArchive, all.Mode);
    }
}
