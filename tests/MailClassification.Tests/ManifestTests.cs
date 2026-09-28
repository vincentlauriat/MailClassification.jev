using System.Xml.Linq;
using MailClassification.Web;

namespace MailClassification.Tests;

public class ManifestTests
{
    [Fact]
    public void Manifest_is_well_formed_and_points_to_the_https_origin()
    {
        var xml = UiServer.BuildManifest(5178);
        var doc = XDocument.Parse(xml);
        var ns = doc.Root!.Name.Namespace;
        Assert.Equal("OfficeApp", doc.Root.Name.LocalName);
        Assert.Equal(UiServer.AddInId, doc.Root.Element(ns + "Id")!.Value);
        Assert.Equal("Mailbox", doc.Root.Element(ns + "Hosts")!.Element(ns + "Host")!.Attribute("Name")!.Value);
        Assert.Contains("https://localhost:5178/taskpane.html", xml);
        Assert.Contains("https://localhost:5178/icon-80.png", xml);
        Assert.DoesNotContain("http://localhost", xml); // every add-in URL must be HTTPS
        Assert.DoesNotContain("http://127.0.0.1", xml);
        Assert.Contains("ReadWriteItem", xml);
        Assert.DoesNotContain("SupportsPinning", xml); // only valid in VersionOverrides 1.1
    }

    [Fact]
    public void Manifest_carries_the_new_name_and_keeps_the_stable_add_in_id()
    {
        var doc = XDocument.Parse(UiServer.BuildManifest(5178));
        var ns = doc.Root!.Name.Namespace;
        Assert.Equal("7c1f3f0e-6d2a-4b5e-9c1a-2f0e8a5d4b31", doc.Root.Element(ns + "Id")!.Value); // renaming must not change the identity
        Assert.Equal("MailClassification", doc.Root.Element(ns + "DisplayName")!.Attribute("DefaultValue")!.Value);
        Assert.Equal("MailClassification", doc.Root.Element(ns + "ProviderName")!.Value);
        Assert.DoesNotContain("jevOutlook", doc.ToString(), StringComparison.OrdinalIgnoreCase);
    }
}
