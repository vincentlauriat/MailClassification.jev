namespace MailClassification.Triage;

public sealed record ResultRow(string From, string Subject, string Label, string Confidence, string Stage, string Action);

/// <summary>Receives events and per-message results while a session runs.</summary>
public interface ITriageSink
{
    void Event(string level, string message);
    void Result(ResultRow row);
    void BatchCompleted(TriageJob job);
}
