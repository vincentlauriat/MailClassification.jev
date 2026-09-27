using System.Text.RegularExpressions;

namespace JevOutlook.Rules;

public sealed class RuleValidationException(string message) : Exception(message);

/// <summary>
/// Validates user-created classification rules before they are saved or used
/// for a run. Outlook categories are plain strings; the only technical
/// constraints are the separator characters that Outlook clients use in
/// category fields, and control characters.
/// </summary>
public static class RuleValidator
{
    private static readonly Regex IdPattern = new("^[a-zA-Z0-9_-]{1,64}$", RegexOptions.Compiled);
    private static readonly Regex ControlChars = new("[\\x00-\\x1f\\x7f]", RegexOptions.Compiled);
    private static readonly string[] ForbiddenIds = ["__proto__", "constructor", "prototype"];

    public static List<LabelRule> ValidateAndNormalize(IEnumerable<LabelRule?>? rules)
    {
        if (rules is null) throw new RuleValidationException("Label rules must be an array.");
        var list = rules.ToList();
        if (list.Count == 0) throw new RuleValidationException("Add at least one classification category.");
        if (list.Count > AppConstants.MaxRules)
            throw new RuleValidationException($"Maximum {AppConstants.MaxRules} classification categories.");

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var normalized = new List<LabelRule>(list.Count);

        for (var index = 0; index < list.Count; index++)
        {
            var rule = list[index];
            var name = (rule?.Name ?? string.Empty).Trim();
            var description = (rule?.Description ?? string.Empty).Trim();
            var spam = rule?.Spam ?? false;

            if (name.Length == 0) throw new RuleValidationException($"Category #{index + 1} has no name.");
            if (name.Length > AppConstants.MaxRuleName) throw new RuleValidationException($"Category \"{name}\" is too long.");
            if (description.Length == 0) throw new RuleValidationException($"Add classification criteria for category \"{name}\".");
            if (description.Length > AppConstants.MaxRuleDescription)
                throw new RuleValidationException($"Description for \"{name}\" is too long.");

            if (ControlChars.IsMatch(name))
                throw new RuleValidationException($"Category \"{name}\" contains control characters.");
            if (name.Contains(',') || name.Contains(';'))
                throw new RuleValidationException($"Category \"{name}\" must not contain ',' or ';' — Outlook uses them as category separators.");

            if (!names.Add(name)) throw new RuleValidationException($"Duplicate category name: {name}");

            var id = (rule?.Id ?? string.Empty).Trim();
            if (!IdPattern.IsMatch(id) || ForbiddenIds.Contains(id) || ids.Contains(id))
            {
                id = "rule-" + Guid.NewGuid().ToString("N")[..8];
            }
            ids.Add(id);

            normalized.Add(new LabelRule(id, name, description, spam));
        }

        return normalized;
    }

    public static LabelRule? ById(IReadOnlyList<LabelRule> rules, string? id) =>
        id is null ? null : rules.FirstOrDefault(r => r.Id == id);

    public static string NameById(IReadOnlyList<LabelRule> rules, string? id) =>
        ById(rules, id)?.Name ?? "Not assigned";

    /// <summary>The safe non-archive fallback used when a body cannot be extracted.</summary>
    public static LabelRule? ReviewFallback(IReadOnlyList<LabelRule> rules) =>
        rules.FirstOrDefault(r => !r.Spam &&
            (r.Id == "review" || string.Equals(r.Name, "review", StringComparison.OrdinalIgnoreCase)));
}
