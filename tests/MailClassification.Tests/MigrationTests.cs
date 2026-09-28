using MailClassification.Storage;

namespace MailClassification.Tests;

public class MigrationTests : IDisposable
{
    private readonly string _home = Path.Combine(Path.GetTempPath(), "mailclassification-tests", Guid.NewGuid().ToString("N"));
    private string Legacy => Path.Combine(_home, ".jevoutlook");
    private string Current => Path.Combine(_home, ".mailclassification");

    public MigrationTests() => Directory.CreateDirectory(_home);

    public void Dispose() => Directory.Delete(_home, recursive: true);

    private void SeedLegacy()
    {
        Directory.CreateDirectory(Path.Combine(Legacy, "accounts", "work"));
        Directory.CreateDirectory(Path.Combine(Legacy, "bin"));
        File.WriteAllText(Path.Combine(Legacy, "config.json"), "{\"clientId\":\"x\"}");
        File.WriteAllText(Path.Combine(Legacy, "accounts", "work", "auth-record.json"), "{}");
        File.WriteAllText(Path.Combine(Legacy, "bin", "jevoutlook"), "old exe");
    }

    [Fact]
    public void Legacy_root_is_moved_when_the_new_root_does_not_exist()
    {
        SeedLegacy();
        var notice = AppPaths.MigrateRoot(Legacy, Current);
        Assert.NotNull(notice);
        Assert.False(Directory.Exists(Legacy));
        Assert.Equal("{\"clientId\":\"x\"}", File.ReadAllText(Path.Combine(Current, "config.json")));
        Assert.True(File.Exists(Path.Combine(Current, "accounts", "work", "auth-record.json")));
    }

    [Fact]
    public void Migration_is_idempotent()
    {
        SeedLegacy();
        Assert.NotNull(AppPaths.MigrateRoot(Legacy, Current));
        Assert.Null(AppPaths.MigrateRoot(Legacy, Current));
        Assert.True(File.Exists(Path.Combine(Current, "config.json")));
    }

    [Fact]
    public void Nothing_happens_without_a_legacy_root()
    {
        Assert.Null(AppPaths.MigrateRoot(Legacy, Current));
        Assert.False(Directory.Exists(Current));
    }

    [Fact]
    public void New_root_with_state_is_never_overwritten()
    {
        SeedLegacy();
        Directory.CreateDirectory(Current);
        File.WriteAllText(Path.Combine(Current, "config.json"), "{\"clientId\":\"new\"}");
        Assert.Null(AppPaths.MigrateRoot(Legacy, Current));
        Assert.Equal("{\"clientId\":\"new\"}", File.ReadAllText(Path.Combine(Current, "config.json")));
        Assert.True(File.Exists(Path.Combine(Legacy, "config.json"))); // left untouched
    }

    [Fact]
    public void State_merges_into_a_new_root_that_only_holds_the_published_bin()
    {
        SeedLegacy();
        Directory.CreateDirectory(Path.Combine(Current, "bin"));
        File.WriteAllText(Path.Combine(Current, "bin", "mailclassification"), "new exe");

        var notice = AppPaths.MigrateRoot(Legacy, Current);

        Assert.NotNull(notice);
        Assert.True(File.Exists(Path.Combine(Current, "config.json")));
        Assert.True(File.Exists(Path.Combine(Current, "accounts", "work", "auth-record.json")));
        Assert.Equal("new exe", File.ReadAllText(Path.Combine(Current, "bin", "mailclassification")));
        Assert.False(File.Exists(Path.Combine(Current, "bin", "jevoutlook")));
        // The old agent's executable stays where its plist points until 'service install' replaces it.
        Assert.Equal("old exe", File.ReadAllText(Path.Combine(Legacy, "bin", "jevoutlook")));
        Assert.Equal(["bin"], Directory.EnumerateFileSystemEntries(Legacy).Select(p => Path.GetFileName(p)).ToArray());
        Assert.Null(AppPaths.MigrateRoot(Legacy, Current)); // second run: no-op
    }

    [Fact]
    public void Legacy_root_without_bin_is_removed_after_a_merge()
    {
        Directory.CreateDirectory(Legacy);
        File.WriteAllText(Path.Combine(Legacy, "rules.json"), "[]");
        Directory.CreateDirectory(Path.Combine(Current, "bin"));
        Assert.NotNull(AppPaths.MigrateRoot(Legacy, Current));
        Assert.False(Directory.Exists(Legacy));
        Assert.True(File.Exists(Path.Combine(Current, "rules.json")));
    }

    [Fact]
    public void Keychain_read_hits_the_current_service_first()
    {
        var stored = new List<string>();
        var deleted = 0;
        var value = SecretStore.GetWithLegacyFallback(
            service => service == "mailclassification" ? "new-secret" : "old-secret", stored.Add, () => deleted++);
        Assert.Equal("new-secret", value);
        Assert.Empty(stored);
        Assert.Equal(0, deleted);
    }

    [Fact]
    public void Keychain_legacy_item_is_copied_then_deleted_on_a_miss()
    {
        var stored = new List<string>();
        var deleted = 0;
        var value = SecretStore.GetWithLegacyFallback(
            service => service == "jevoutlook" ? "old-secret" : null, stored.Add, () => deleted++);
        Assert.Equal("old-secret", value);
        Assert.Equal(["old-secret"], stored);
        Assert.Equal(1, deleted);
    }

    [Fact]
    public void Keychain_legacy_item_is_kept_when_the_copy_fails()
    {
        var deleted = 0;
        var value = SecretStore.GetWithLegacyFallback(
            service => service == "jevoutlook" ? "old-secret" : null,
            _ => throw new InvalidOperationException("keychain locked"), () => deleted++);
        Assert.Equal("old-secret", value);
        Assert.Equal(0, deleted);
    }

    [Fact]
    public void Keychain_miss_everywhere_returns_null()
    {
        Assert.Null(SecretStore.GetWithLegacyFallback(_ => null, _ => throw new Xunit.Sdk.XunitException("no write"), () => throw new Xunit.Sdk.XunitException("no delete")));
    }
}
