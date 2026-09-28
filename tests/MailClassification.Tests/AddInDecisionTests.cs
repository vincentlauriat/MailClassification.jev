using JevOutlook.Web;

namespace JevOutlook.Tests;

public class AddInDecisionTests
{
    [Theory]
    [InlineData(false, 0.74, true)]  // unsure metadata pass
    [InlineData(false, 0.80, false)] // confident enough, not archive-eligible
    [InlineData(true, 0.80, true)]   // archive-eligible without archive confidence: read the body
    [InlineData(true, 0.93, false)]
    [InlineData(true, 0.99, false)]
    public void Full_review_follows_the_batch_escalation_rule(bool archiveEligible, double confidence, bool expected)
    {
        Assert.Equal(expected, UiServer.NeedsFullReview(archiveEligible, confidence));
    }

    [Theory]
    [InlineData(false, 0.99, false)] // never archive a category that is not eligible
    [InlineData(true, 0.929, false)]
    [InlineData(true, 0.93, true)]
    [InlineData(true, 0.99, true)]
    public void Archive_is_offered_only_at_the_archive_confidence(bool archiveEligible, double confidence, bool expected)
    {
        Assert.Equal(expected, UiServer.CanArchive(archiveEligible, confidence));
    }
}
