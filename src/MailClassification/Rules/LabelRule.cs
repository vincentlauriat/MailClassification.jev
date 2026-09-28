namespace MailClassification.Rules;

/// <summary>
/// One user-defined classification rule. <see cref="Spam"/> means the category
/// is archive-eligible: it MAY be moved to the Archive folder, only in
/// <c>labels_archive</c> mode and only above the archive confidence threshold.
/// </summary>
public sealed record LabelRule(string Id, string Name, string Description, bool Spam);

public sealed record Playbook(string Id, string Name, string Summary, IReadOnlyList<LabelRule> Rules);
