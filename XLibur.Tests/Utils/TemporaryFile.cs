using System;
using System.IO;

namespace XLibur.Tests.Utils;

internal sealed class TemporaryFile : IDisposable
{
    private bool _disposed;

    internal TemporaryFile()
        : this(System.IO.Path.ChangeExtension(
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName()), "xlsx"))
    {
    }

    internal TemporaryFile(string path)
    {
        Path = path;
    }

    public string Path { get; private set; }

    public void Dispose()
    {
        Dispose(disposing: true);
    }

    private void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
            File.Delete(Path);

        _disposed = true;
    }

    public override string ToString()
    {
        return Path;
    }
}