using System.Globalization;

namespace JevOutlook.Triage;

/// <summary>Validated run configuration (the counterpart of the jevMail run form).</summary>
public sealed record RunOptions(
    string Scope,
    bool UnreadOnly,
    bool DryRun,
    string Mode,
    int? Limit,
    double MaxSpendUsd,
    double MetadataThreshold,
    double ArchiveThreshold)
{
    public static RunOptions Normalize(string? scope, bool unreadOnly, bool dryRun, string? mode, string? limitRaw,
        double maxSpendUsd, double metadataThreshold, double archiveThreshold)
    {
        int? limit;
        var raw = (limitRaw ?? AppConstants.DefaultLimit.ToString(CultureInfo.InvariantCulture)).Trim();
        if (string.Equals(raw, "all", StringComparison.OrdinalIgnoreCase)) limit = null;
        else if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0 && parsed <= 100_000) limit = parsed;
        else throw new ArgumentException("Message limit must be a positive number (e.g. 10, 100, 1000) or 'all'.");

        if (double.IsNaN(maxSpendUsd) || maxSpendUsd <= 0 || maxSpendUsd > 1000)
            throw new ArgumentException("The maximum processing cost must be greater than $0 and no more than $1000.");
        if (double.IsNaN(metadataThreshold) || metadataThreshold < 0.5 || metadataThreshold > 0.99)
            throw new ArgumentException("Metadata confidence threshold must be between 0.50 and 0.99.");
        if (double.IsNaN(archiveThreshold) || archiveThreshold < 0.5 || archiveThreshold > 0.999)
            throw new ArgumentException("Archive confidence threshold must be between 0.50 and 0.999.");
        if (archiveThreshold < metadataThreshold)
            throw new ArgumentException("Archiving confidence must be at least as high as metadata confidence.");

        var normalizedMode = string.Equals(mode, "labels_archive", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(mode, "labels-archive", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(mode, "archive", StringComparison.OrdinalIgnoreCase)
            ? RunMode.LabelsArchive
            : RunMode.Labels;
        var normalizedScope = string.Equals(scope, "all", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(scope, "allmail", StringComparison.OrdinalIgnoreCase)
            ? "all"
            : "inbox";

        return new RunOptions(normalizedScope, unreadOnly, dryRun, normalizedMode, limit, maxSpendUsd, metadataThreshold, archiveThreshold);
    }
}
