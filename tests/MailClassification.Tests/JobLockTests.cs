using JevOutlook.Storage;

namespace JevOutlook.Tests;

public class JobLockTests
{
    private static string TempLockPath() =>
        Path.Combine(Path.GetTempPath(), "jevoutlook-tests", Guid.NewGuid().ToString("N"), "job.lock");

    [Fact]
    public void Second_acquire_fails_while_held_even_in_the_same_process()
    {
        var path = TempLockPath();
        using (var first = JobStore.TryLockFile(path))
        {
            Assert.NotNull(first);
            Assert.Null(JobStore.TryLockFile(path));
        }
        using var again = JobStore.TryLockFile(path);
        Assert.NotNull(again);
        Assert.True(File.Exists(path)); // the lock file is kept: unlinking a flock file would allow two holders
    }

    [Fact]
    public void Lock_is_released_when_the_holder_throws()
    {
        var path = TempLockPath();
        void FailingBatch()
        {
            using var held = JobStore.TryLockFile(path) ?? throw new Xunit.Sdk.XunitException("not acquired");
            Assert.Null(JobStore.TryLockFile(path));
            throw new InvalidOperationException("batch failed");
        }
        Assert.Throws<InvalidOperationException>(FailingBatch);
        using var after = JobStore.TryLockFile(path);
        Assert.NotNull(after);
    }

    [Fact]
    public void Lock_file_is_private_even_when_an_older_build_created_it_world_readable()
    {
        if (OperatingSystem.IsWindows()) return;
        var path = TempLockPath();
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, string.Empty);
        File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
        using var held = JobStore.TryLockFile(path);
        Assert.NotNull(held);
        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(path));
    }
}
