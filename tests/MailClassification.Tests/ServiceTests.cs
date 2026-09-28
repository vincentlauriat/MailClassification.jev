using System.Diagnostics;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using MailClassification.Cli;
using MailClassification.Web;

namespace MailClassification.Tests;

public class ServiceTests
{
    private static readonly string[] UiArgs = ["ui", "--no-open", "--port", "5177", "--https-port", "5178"];

    private static XDocument Parse(string xml) =>
        XDocument.Load(XmlReader.Create(new StringReader(xml), new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }));

    /// <summary>Plist dict as key → value element (keys and values alternate).</summary>
    private static Dictionary<string, XElement> Dict(XElement dict)
    {
        var items = dict.Elements().ToList();
        var map = new Dictionary<string, XElement>();
        for (var i = 0; i + 1 < items.Count; i += 2) map[items[i].Value] = items[i + 1];
        return map;
    }

    [Fact]
    public void Plist_carries_label_arguments_keepalive_and_log_path()
    {
        var xml = ServiceCommand.BuildLaunchAgentPlist("/Users/me/.mailclassification/bin/mailclassification", UiArgs,
            new Dictionary<string, string> { ["DOTNET_ROOT"] = "/usr/local/share/dotnet" }, "/Users/me/.mailclassification/logs/ui.log");
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"UTF-8\"?>", xml);
        Assert.Contains("<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\"", xml);

        var doc = Parse(xml);
        Assert.Equal("plist", doc.Root!.Name.LocalName);
        var d = Dict(doc.Root.Element("dict")!);
        Assert.Equal("com.vincentlauriat.mailclassification", d["Label"].Value);
        Assert.Equal(["/Users/me/.mailclassification/bin/mailclassification", .. UiArgs], d["ProgramArguments"].Elements("string").Select(e => e.Value).ToArray());
        Assert.Equal("true", d["RunAtLoad"].Name.LocalName);
        Assert.Equal("true", d["KeepAlive"].Name.LocalName);
        Assert.Equal("30", d["ThrottleInterval"].Value);
        Assert.Equal("/Users/me/.mailclassification/logs/ui.log", d["StandardOutPath"].Value);
        Assert.Equal("/Users/me/.mailclassification/logs/ui.log", d["StandardErrorPath"].Value);
        Assert.Equal("/usr/local/share/dotnet", Dict(d["EnvironmentVariables"])["DOTNET_ROOT"].Value);
    }

    [Fact]
    public void Plist_never_contains_an_api_key()
    {
        var env = new Dictionary<string, string>
        {
            ["OPENROUTER_API_KEY"] = "sk-or-secret",
            ["JEV_API_KEY"] = "jev-secret",
            ["MAILCLASSIFICATION_HOME"] = "/tmp/jev",
            ["JEVOUTLOOK_HOME"] = "/tmp/legacy",
            ["PATH"] = "/usr/bin",
        };
        var xml = ServiceCommand.BuildLaunchAgentPlist("/opt/mailclassification", UiArgs, env, "/tmp/jev/logs/ui.log");
        Assert.DoesNotContain("OPENROUTER_API_KEY", xml);
        Assert.DoesNotContain("JEV_API_KEY", xml);
        Assert.DoesNotContain("secret", xml);
        Assert.DoesNotContain("PATH", xml);
        var vars = Dict(Dict(Parse(xml).Root!.Element("dict")!)["EnvironmentVariables"]);
        Assert.Equal(["JEVOUTLOOK_HOME", "MAILCLASSIFICATION_HOME"], vars.Keys.ToArray());
    }

    [Fact]
    public void Plist_omits_environment_when_nothing_is_allowed()
    {
        var xml = ServiceCommand.BuildLaunchAgentPlist("/opt/mailclassification", UiArgs, new Dictionary<string, string> { ["OPENROUTER_API_KEY"] = "x" }, "/tmp/ui.log");
        Assert.DoesNotContain("EnvironmentVariables", xml);
    }

    [Fact]
    public void Plist_escapes_paths_with_spaces_and_ampersands()
    {
        const string exe = "/Users/me/My Apps & Tools/<jev>/mailclassification";
        const string log = "/Users/me/Logs & Co/ui.log";
        var xml = ServiceCommand.BuildLaunchAgentPlist(exe, UiArgs, new Dictionary<string, string>(), log);
        Assert.Contains("My Apps &amp; Tools/&lt;jev&gt;/mailclassification", xml);
        var d = Dict(Parse(xml).Root!.Element("dict")!);
        Assert.Equal(exe, d["ProgramArguments"].Elements("string").First().Value);
        Assert.Equal(log, d["StandardOutPath"].Value);
    }

    [Fact]
    public void Plist_passes_plutil_lint_on_macos()
    {
        if (!OperatingSystem.IsMacOS() || !File.Exists("/usr/bin/plutil")) return;
        var file = Path.Combine(Path.GetTempPath(), $"mailclassification-test-{Guid.NewGuid():N}.plist");
        try
        {
            File.WriteAllText(file, ServiceCommand.BuildLaunchAgentPlist("/Users/me/A & B/mailclassification", UiArgs,
                new Dictionary<string, string> { ["MAILCLASSIFICATION_HOME"] = "/tmp/jev" }, "/tmp/ui.log"), new UTF8Encoding(false));
            var psi = new ProcessStartInfo("/usr/bin/plutil") { RedirectStandardOutput = true, RedirectStandardError = true };
            psi.ArgumentList.Add("-lint");
            psi.ArgumentList.Add(file);
            using var p = Process.Start(psi)!;
            var output = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
            p.WaitForExit();
            Assert.True(p.ExitCode == 0, output);
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Fact]
    public void Allowlist_covers_http_port_only_without_https()
    {
        var (hosts, origins) = UiServer.AllowedOrigins(5177, null);
        Assert.Equal(new HashSet<string> { "127.0.0.1:5177", "localhost:5177" }, hosts);
        Assert.Equal(new HashSet<string> { "http://127.0.0.1:5177", "http://localhost:5177" }, origins);
        Assert.Contains("LOCALHOST:5177", hosts); // host names are case-insensitive
        Assert.DoesNotContain("evil.example:5177", hosts);
    }

    [Fact]
    public void Allowlist_adds_the_https_port_when_enabled()
    {
        var (hosts, origins) = UiServer.AllowedOrigins(5177, 5178);
        Assert.Equal(new HashSet<string> { "127.0.0.1:5177", "localhost:5177", "localhost:5178", "127.0.0.1:5178" }, hosts);
        Assert.Equal(new HashSet<string> { "http://127.0.0.1:5177", "http://localhost:5177", "https://localhost:5178", "https://127.0.0.1:5178" }, origins);
        Assert.DoesNotContain("http://localhost:5178", origins);
    }

    [Theory]
    [InlineData("{\"app\":\"mailclassification\",\"version\":\"0.1.0\",\"https\":true}", "0.1.0")]
    [InlineData("{\"app\":\"mailclassification\"}", "")]
    [InlineData("{\"app\":\"jevoutlook\",\"version\":\"0.1.0\",\"https\":true}", "0.1.0")] // pre-rename server
    [InlineData("{\"app\":\"JevOutlook\",\"version\":\"0.1.0\"}", null)]
    [InlineData("{\"app\":\"other\",\"version\":\"1.0\"}", null)]
    [InlineData("<html>Not found</html>", null)]
    [InlineData("", null)]
    [InlineData("[1,2]", null)]
    public void Health_body_is_recognised_only_for_mailclassification_or_legacy_jevoutlook(string body, string? expected)
    {
        Assert.Equal(expected, UiServer.ParseHealth(body));
    }

    [Fact]
    public void Legacy_launch_agent_is_detected_only_when_its_plist_exists()
    {
        var dir = Path.Combine(Path.GetTempPath(), "mailclassification-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            Assert.Null(ServiceCommand.LegacyPlistIn(dir));
            File.WriteAllText(Path.Combine(dir, ServiceCommand.Label + ".plist"), "<plist/>"); // the current agent is not "legacy"
            Assert.Null(ServiceCommand.LegacyPlistIn(dir));
            var legacy = Path.Combine(dir, "com.vincentlauriat.jevoutlook.plist");
            File.WriteAllText(legacy, "<plist/>");
            Assert.Equal(legacy, ServiceCommand.LegacyPlistIn(dir));
            Assert.Null(ServiceCommand.LegacyPlistIn(Path.Combine(dir, "missing")));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }
}
