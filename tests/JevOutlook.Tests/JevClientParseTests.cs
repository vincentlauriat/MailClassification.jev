using JevOutlook;
using JevOutlook.Jev;
using JevOutlook.Rules;

namespace JevOutlook.Tests;

public class JevClientParseTests
{
    private static readonly IReadOnlyList<LabelRule> Rules =
    [
        new("a", "alpha", "d", false),
        new("b", "beta", "d", true),
        new("c", "gamma", "d", false),
    ];

    private static string Ok(string choice = "L1", string probs = "{\"L0\":0.1,\"L1\":0.8,\"L2\":0.1}", string confidence = "0.72", string usage = "{\"input_tokens\":300,\"cost\":0.00001}") =>
        $"{{\"model\":\"jev-1.13.0\",\"provider\":\"TypeSafe\",\"answers\":{{\"label\":{{\"type\":\"choice\",\"choice\":\"{choice}\",\"probabilities\":{probs},\"confidence\":{confidence}}}}},\"usage\":{usage}}}";

    [Fact]
    public void Valid_response_maps_choice_confidence_and_reported_cost()
    {
        var result = JevClient.Parse(200, 0, Ok(), Rules, 0.5);
        Assert.True(result.Ok);
        Assert.Equal("b", result.RuleId);
        Assert.Equal(0.72, result.Confidence);
        Assert.Equal(0.00001, result.CostUsd, 10);
        Assert.Equal("reported", result.CostSource);
        Assert.Equal(300, result.InputTokens);
        Assert.Equal("L1", result.ProbabilityArgmax);
        Assert.False(result.ChoiceDiffersFromArgmax);
        Assert.Equal("jev-1.13.0", result.Model);
    }

    [Fact]
    public void Cost_falls_back_to_input_tokens_then_estimate()
    {
        var tokens = JevClient.Parse(200, 0, Ok(usage: "{\"input_tokens\":1000000}"), Rules, 0.5);
        Assert.Equal("input_tokens", tokens.CostSource);
        Assert.Equal(AppConstants.InputRateUsdPerMillion, tokens.CostUsd, 9);

        var estimated = JevClient.Parse(200, 0, Ok(usage: "{}"), Rules, 0.0042);
        Assert.Equal("estimated", estimated.CostSource);
        Assert.Equal(0.0042, estimated.CostUsd, 9);
    }

    [Theory]
    [InlineData(401, "invalid_api_key", false)]
    [InlineData(402, "insufficient_credits", false)]
    [InlineData(403, "access_denied", false)]
    [InlineData(429, "provider_temporarily_unavailable", true)]
    [InlineData(503, "provider_temporarily_unavailable", true)]
    [InlineData(400, "request_rejected", false)]
    public void Http_errors_are_session_scoped(int status, string code, bool retryable)
    {
        var result = JevClient.Parse(status, 1500, "{}", Rules, 0.001);
        Assert.False(result.Ok);
        Assert.Equal(code, result.Code);
        Assert.Equal(retryable, result.Retryable);
        Assert.Equal("session", result.FailureScope);
        Assert.Equal(1500, result.RetryAfterMs);
        // A 429 is rejected before inference (not billed); everything else keeps the estimate reserved.
        Assert.Equal(status == 429 ? 0 : 0.001, result.CostUsd, 9);
    }

    [Fact]
    public void Retry_after_is_capped()
    {
        var result = JevClient.Parse(429, 60_000, "{}", Rules, 0);
        Assert.Equal(AppConstants.MaxJevRetryDelayMs, result.RetryAfterMs);
    }

    [Theory]
    [InlineData("not json", "invalid_json")]
    [InlineData("{\"answers\":{}}", "missing_answer")]
    public void Malformed_bodies_are_message_scoped(string body, string code)
    {
        var result = JevClient.Parse(200, 0, body, Rules, 0);
        Assert.False(result.Ok);
        Assert.Equal(code, result.Code);
        Assert.Equal("message", result.FailureScope);
        Assert.True(result.Retryable);
    }

    [Fact]
    public void Unknown_choice_is_rejected()
    {
        Assert.Equal("invalid_choice", JevClient.Parse(200, 0, Ok(choice: "L7"), Rules, 0).Code);
        Assert.Equal("invalid_choice", JevClient.Parse(200, 0, Ok(choice: "L01"), Rules, 0).Code);
        Assert.Equal("invalid_choice", JevClient.Parse(200, 0, Ok(choice: "beta"), Rules, 0).Code);
    }

    [Fact]
    public void Confidence_must_be_in_range()
    {
        Assert.Equal("invalid_confidence", JevClient.Parse(200, 0, Ok(confidence: "1.5"), Rules, 0).Code);
        Assert.Equal("invalid_confidence", JevClient.Parse(200, 0, Ok(confidence: "\"high\""), Rules, 0).Code);
    }

    [Fact]
    public void Probability_contract_is_strict()
    {
        Assert.Equal("missing_probabilities", JevClient.Parse(200, 0, Ok(probs: "{\"L0\":0.2,\"L1\":0.8}"), Rules, 0).Code);
        Assert.Equal("unexpected_probability_keys", JevClient.Parse(200, 0, Ok(probs: "{\"L0\":0.1,\"L1\":0.8,\"L2\":0.1,\"L3\":0}"), Rules, 0).Code);
        Assert.Equal("invalid_probability_value", JevClient.Parse(200, 0, Ok(probs: "{\"L0\":-0.1,\"L1\":1.0,\"L2\":0.1}"), Rules, 0).Code);
        Assert.Equal("probability_sum_mismatch", JevClient.Parse(200, 0, Ok(probs: "{\"L0\":0.5,\"L1\":0.5,\"L2\":0.5}"), Rules, 0).Code);
        Assert.Equal("invalid_probabilities", JevClient.Parse(200, 0, Ok(probs: "[0.1,0.8,0.1]"), Rules, 0).Code);
    }

    [Fact]
    public void Floating_point_noise_in_sum_is_tolerated()
    {
        var result = JevClient.Parse(200, 0, Ok(probs: "{\"L0\":0.333,\"L1\":0.334,\"L2\":0.333}"), Rules, 0);
        Assert.True(result.Ok);
        Assert.Equal("L1", result.ProbabilityArgmax);
    }

    [Fact]
    public void Choice_differing_from_argmax_is_flagged_not_rejected()
    {
        var result = JevClient.Parse(200, 0, Ok(choice: "L0", probs: "{\"L0\":0.2,\"L1\":0.7,\"L2\":0.1}"), Rules, 0);
        Assert.True(result.Ok);
        Assert.Equal("a", result.RuleId);
        Assert.True(result.ChoiceDiffersFromArgmax);
    }
}
