using JevOutlook;
using JevOutlook.Rules;

namespace JevOutlook.Tests;

public class RuleValidatorTests
{
    [Fact]
    public void Default_playbook_is_valid_and_keeps_ids()
    {
        var normalized = RuleValidator.ValidateAndNormalize(Playbooks.DefaultRules);
        Assert.Equal(Playbooks.DefaultRules.Count, normalized.Count);
        Assert.Equal(Playbooks.DefaultRules.Select(r => r.Id), normalized.Select(r => r.Id));
        Assert.Contains(normalized, r => r.Id == "review" && !r.Spam);
    }

    [Fact]
    public void Every_playbook_is_valid()
    {
        foreach (var playbook in Playbooks.All)
        {
            var rules = RuleValidator.ValidateAndNormalize(playbook.Rules);
            Assert.NotEmpty(rules);
            Assert.NotNull(RuleValidator.ReviewFallback(rules));
        }
    }

    [Fact]
    public void Rejects_empty_and_too_many()
    {
        Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize([]));
        var tooMany = Enumerable.Range(0, AppConstants.MaxRules + 1).Select(i => new LabelRule("r" + i, "name" + i, "desc", false));
        Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize(tooMany));
    }

    [Fact]
    public void Rejects_separators_and_duplicates()
    {
        Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize([new LabelRule("a", "a,b", "d", false)]));
        Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize([new LabelRule("a", "a;b", "d", false)]));
        Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize([new LabelRule("a", "Same", "d", false), new LabelRule("b", "same", "d", false)]));
        Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize([new LabelRule("a", "ok", "", false)]));
    }

    [Theory]
    [InlineData("A b", "A_b")]
    [InlineData("À traiter", "__traiter")]
    [InlineData("junk", "jev-junk")]
    [InlineData("$Junk", "Jev-junk")]
    public void Rejects_names_stored_as_the_same_imap_keyword(string first, string second)
    {
        var ex = Assert.Throws<RuleValidationException>(() => RuleValidator.ValidateAndNormalize(
            [new LabelRule("a", first, "d", false), new LabelRule("b", second, "d", false)]));
        Assert.Contains("IMAP keyword", ex.Message);
        Assert.Contains(first, ex.Message);
        Assert.Contains(second, ex.Message);
    }

    [Fact]
    public void Regenerates_invalid_or_duplicate_ids_and_trims()
    {
        var rules = RuleValidator.ValidateAndNormalize(
        [
            new LabelRule("dup", "  one ", " first ", true),
            new LabelRule("dup", "two", "second", false),
            new LabelRule("bad id!", "three", "third", false),
            new LabelRule("__proto__", "four", "fourth", false),
        ]);
        Assert.Equal("one", rules[0].Name);
        Assert.Equal("first", rules[0].Description);
        Assert.True(rules[0].Spam);
        Assert.Equal("dup", rules[0].Id);
        Assert.StartsWith("rule-", rules[1].Id);
        Assert.StartsWith("rule-", rules[2].Id);
        Assert.StartsWith("rule-", rules[3].Id);
        Assert.Equal(4, rules.Select(r => r.Id).Distinct().Count());
    }
}
