using System;
using System.IO;

namespace XLibur.Tests.Utils;

internal sealed class TemporaryFile : IDisposable
{
    private bool _disposed;

    internal TemporaryFile()
        : this(System.IO.Path.ChangeExtension(System.IO.Path.GetRandomFileName(), "xlsx"))
    {
    }

    internal TemporaryFile(string path)
        : this(path, false)
    {
    }

    private TemporaryFile(string path, bool preserve)
    {
        Path = path;
        Preserve = preserve;
    }

    public string Path { get; private set; }

    private bool Preserve { get; set; }

    public void Dispose()
    {
        Dispose(disposing: true);
    }

    private void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing && !Preserve)
            File.Delete(Path);

        _disposed = true;
    }

    public override string ToString()
    {
        return Path;
    }
}
