using JevOutlook.Graph;
using JevOutlook.Triage;

namespace JevOutlook.Tests;

public class CursorPagingTests
{
    private static readonly DateTimeOffset T0 = new(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

    private static MessageRef Ref(string id, int secondsAgo, params string[] categories) =>
        new(id, T0.AddSeconds(-secondsAgo), categories, "inbox", false);

    private static TriageJob NewJob() => new() { Cursor = T0.AddMinutes(5), CursorBoundaryIds = [] };

    [Fact]
    public void Empty_page_marks_exhaustion()
    {
        var job = NewJob();
        var collected = new List<PendingItem>();
        TriageEngine.ExaminePage(job, new ListPage([], false), 10, _ => true, collected);
        Assert.True(job.Exhausted);
        Assert.Empty(collected);
    }

    [Fact]
    public void Cursor_only_advances_over_examined_items()
    {
        var job = NewJob();
        var collected = new List<PendingItem>();
        var page = new ListPage([Ref("a", 1), Ref("b", 2), Ref("c", 3), Ref("d", 4)], true);
        TriageEngine.ExaminePage(job, page, 2, _ => true, collected);

        Assert.Equal(["a", "b"], collected.Select(c => c.Id));
        Assert.Equal(T0.AddSeconds(-2), job.Cursor);           // stopped after "b"
        Assert.Equal(["b"], job.CursorBoundaryIds);
        Assert.False(job.Exhausted);                            // "c" and "d" are still ahead
    }

    [Fact]
    public void Skipped_items_advance_the_cursor_but_are_not_collected()
    {
        var job = NewJob();
        var collected = new List<PendingItem>();
        var page = new ListPage([Ref("a", 1, "jev-triaged"), Ref("b", 2), Ref("c", 3)], false);
        TriageEngine.ExaminePage(job, page, 10, r => !r.Categories.Contains("jev-triaged"), collected);

        Assert.Equal(["b", "c"], collected.Select(c => c.Id));
        Assert.Equal(T0.AddSeconds(-3), job.Cursor);
        Assert.True(job.Exhausted);                             // short page fully examined
    }

    [Fact]
    public void Ties_at_the_boundary_are_neither_repeated_nor_lost()
    {
        var job = NewJob();
        var collected = new List<PendingItem>();
        // Two messages share the same timestamp; only one fits in this batch.
        TriageEngine.ExaminePage(job, new ListPage([Ref("a", 1), Ref("b", 5), Ref("c", 5), Ref("d", 9)], true), 2, _ => true, collected);
        Assert.Equal(["a", "b"], collected.Select(c => c.Id));
        Assert.Equal(T0.AddSeconds(-5), job.Cursor);
        Assert.Equal(["b"], job.CursorBoundaryIds);

        // Next page starts at the cursor (le) and returns the tie again.
        collected.Clear();
        TriageEngine.ExaminePage(job, new ListPage([Ref("b", 5), Ref("c", 5), Ref("d", 9)], false), 2, _ => true, collected);
        Assert.Equal(["c", "d"], collected.Select(c => c.Id));
        Assert.Equal(T0.AddSeconds(-9), job.Cursor);
        Assert.True(job.Exhausted);
    }

    [Fact]
    public void Boundary_grows_when_several_items_share_the_cursor_timestamp()
    {
        var job = NewJob();
        var collected = new List<PendingItem>();
        TriageEngine.ExaminePage(job, new ListPage([Ref("a", 5), Ref("b", 5), Ref("c", 5)], false), 10, _ => true, collected);
        Assert.Equal(["a", "b", "c"], job.CursorBoundaryIds);
        Assert.Equal(T0.AddSeconds(-5), job.Cursor);
    }

    [Fact]
    public void Full_page_of_boundary_ids_switches_to_an_exclusive_cursor()
    {
        var job = NewJob();
        job.Cursor = T0;
        job.CursorBoundaryIds = ["a", "b"];
        var collected = new List<PendingItem>();
        TriageEngine.ExaminePage(job, new ListPage([Ref("a", 0), Ref("b", 0)], true), 10, _ => true, collected);
        Assert.Equal(T0, job.Cursor);
        Assert.True(job.CursorExclusive);
        Assert.Empty(job.CursorBoundaryIds);
        Assert.False(job.Exhausted);
        Assert.Empty(collected);

        // The next (strictly older) page resets the cursor to inclusive mode.
        TriageEngine.ExaminePage(job, new ListPage([Ref("c", 1)], false), 10, _ => true, collected);
        Assert.Equal(["c"], collected.Select(c => c.Id));
        Assert.False(job.CursorExclusive);
        Assert.Equal(["c"], job.CursorBoundaryIds);
    }

    [Fact]
    public void Duplicate_ids_are_never_collected_twice()
    {
        var job = NewJob();
        var collected = new List<PendingItem> { new() { Id = "a" } };
        TriageEngine.ExaminePage(job, new ListPage([Ref("a", 1), Ref("b", 2)], false), 10, _ => true, collected);
        Assert.Equal(["a", "b"], collected.Select(c => c.Id));
    }
}
