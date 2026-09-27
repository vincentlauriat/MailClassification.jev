using JevOutlook.Rules;

namespace JevOutlook.Storage;

/// <summary>Saved rule set (falls back to the Universal inbox playbook).</summary>
public static class RuleStore
{
    public static List<LabelRule> Load()
    {
        var saved = JsonStore.Load<List<LabelRule>>(AppPaths.Rules);
        if (saved is null) return Playbooks.DefaultRules.ToList();
        try
        {
            return RuleValidator.ValidateAndNormalize(saved);
        }
        catch (RuleValidationException)
        {
            return Playbooks.DefaultRules.ToList();
        }
    }

    public static List<LabelRule> Save(IEnumerable<LabelRule> rules)
    {
        var normalized = RuleValidator.ValidateAndNormalize(rules);
        JsonStore.Save(AppPaths.Rules, normalized);
        return normalized;
    }

    public static List<LabelRule> Reset() => Save(Playbooks.DefaultRules);
}
