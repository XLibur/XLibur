using System.Diagnostics;
using XLibur.Operations;

namespace XLibur.Operations.Tests;

internal sealed class TempDirectory : IDisposable
{
    public TempDirectory(string? parent = null)
    {
        Path = System.IO.Path.Combine(parent ?? System.IO.Path.GetTempPath(), "xlibur-ops-" + Guid.NewGuid().ToString("N")[..12]);
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string File(string name, string content = "x")
    {
        var path = System.IO.Path.Combine(Path, name);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        System.IO.File.WriteAllText(path, content);
        return path;
    }

    public string Combine(params string[] parts) => System.IO.Path.Combine([Path, .. parts]);

    public void Dispose()
    {
        try
        {
            foreach (var link in Directory.EnumerateFileSystemEntries(Path, "*", SearchOption.AllDirectories)
                         .Where(p => new FileInfo(p).LinkTarget is not null).ToList())
            {
                if (Directory.Exists(link))
                    Directory.Delete(link);
                else
                    System.IO.File.Delete(link);
            }

            Directory.Delete(Path, recursive: true);
        }
        catch (IOException)
        {
        }
    }
}

internal static class Expect
{
    public static async Task Refused(Action action, ErrorKind kind, string code)
    {
        OperationException? caught = null;
        try
        {
            action();
        }
        catch (OperationException e)
        {
            caught = e;
        }

        await Assert.That(caught).IsNotNull();
        await Assert.That(caught!.Error.Code).IsEqualTo(code);
        await Assert.That(caught.Error.Kind).IsEqualTo(kind);
    }
}

internal static class Junctions
{
    /// <summary>Creates a directory junction, which Windows allows without elevation.</summary>
    public static void Create(string link, string target)
    {
        var process = Process.Start(new ProcessStartInfo("cmd.exe", ["/c", "mklink", "/J", link, target])
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        })!;
        process.WaitForExit();
        if (process.ExitCode != 0)
            throw new InvalidOperationException("mklink /J failed: " + process.StandardError.ReadToEnd());
    }

    public static void CreateSymbolicDirectoryOrSkip(string link, string target)
    {
        try
        {
            Directory.CreateSymbolicLink(link, target);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            Skip.Test($"Creating a symbolic link needs Developer Mode or elevation here: {e.Message}");
        }
    }

    public static void CreateSymbolicFileOrSkip(string link, string target)
    {
        try
        {
            File.CreateSymbolicLink(link, target);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            Skip.Test($"Creating a symbolic link needs Developer Mode or elevation here: {e.Message}");
        }
    }
}
