using System.Xml.Linq;
using JevOutlook.Web;

namespace JevOutlook.Tests;

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
}
