using System.Text.Json;
using System.Text.Json.Nodes;
using JevOutlook;
using JevOutlook.Graph;
using JevOutlook.Jev;
using JevOutlook.Rules;

namespace JevOutlook.Tests;

public class JevPayloadTests
{
    private static readonly IReadOnlyList<LabelRule> Rules = Playbooks.DefaultRules;

    [Fact]
    public void Metadata_payload_has_neutral_option_ids_and_no_body()
    {
        var payload = JevPayloadBuilder.Build(JevPayloadBuilder.SampleMetadata(), Rules, "metadata", "ignored", "~typesafe/jev-latest");
        Assert.Equal("~typesafe/jev-latest", payload["model"]!.GetValue<string>());
        Assert.Equal("metadata_only", payload["state"]!["evidence_stage"]!.GetValue<string>());
        var email = payload["state"]!["email"]!.AsObject();
        Assert.False(email.ContainsKey("body"));
        Assert.False(email.ContainsKey("cc")); // empty fields are dropped
        Assert.Equal("Acme SEO <sales@example.com>", email["from"]!.GetValue<string>());

        var question = payload["questions"]!["label"]!.AsObject();
        Assert.Equal("choice", question["type"]!.GetValue<string>());
        var criteria = question["criteria"]!.AsObject();
        Assert.Equal(Rules.Count, criteria.Count);
        for (var i = 0; i < Rules.Count; i++)
        {
            Assert.Equal(Rules[i].Name, criteria["L" + i]!["outlook_category"]!.GetValue<string>());
            Assert.Equal(Rules[i].Description, criteria["L" + i]!["assign_when"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Full_payload_includes_body_and_full_evidence_policy()
    {
        var payload = JevPayloadBuilder.Build(JevPayloadBuilder.SampleMetadata(), Rules, "full", "Hello body", "m");
        Assert.Equal("full_body", payload["state"]!["evidence_stage"]!.GetValue<string>());
        Assert.Equal("Hello body", payload["state"]!["email"]!["body"]!.GetValue<string>());
        Assert.Contains("body", payload["questions"]!["label"]!["instructions"]!["evidence_policy"]!.GetValue<string>());
    }

    [Fact]
    public void Long_header_fields_are_clipped()
    {
        var metadata = new MessageMetadata { Id = "x", Subject = new string('s', 2000), Snippet = new string('p', 5000) };
        var email = JevPayloadBuilder.Build(metadata, Rules, "metadata", "", "m")["state"]!["email"]!.AsObject();
        Assert.Equal(512, email["subject"]!.GetValue<string>().Length);
        Assert.Equal(1200, email["snippet"]!.GetValue<string>().Length);
    }

    [Fact]
    public void Serialize_rejects_oversized_payloads()
    {
        var metadata = new MessageMetadata { Id = "x", Subject = "s" };
        var payload = JevPayloadBuilder.Build(metadata, Rules, "full", new string('b', AppConstants.MaxPayloadBytes + 10), "m");
        Assert.Throws<InvalidOperationException>(() => JevPayloadBuilder.Serialize(payload));
    }

    [Fact]
    public void Estimate_is_positive_and_proportional()
    {
        var small = JevPayloadBuilder.EstimateCost(new string('a', 1000));
        var large = JevPayloadBuilder.EstimateCost(new string('a', 2000));
        Assert.True(small > 0);
        Assert.Equal(2 * small, large, 12);
    }
}
