using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Win32.SafeHandles;

namespace XLibur.Operations.Paths;

/// <summary>
/// Where a file really is. Windows (GetFinalPathNameByHandle) and Linux (/proc/self/fd) answer for an
/// open handle, which no name trick can change. Other systems have no such query here, so the opened
/// name is resolved again instead.
/// </summary>
internal static partial class NativePaths
{
    internal static string FinalPath(SafeFileHandle handle, string openedPath)
    {
        if (OperatingSystem.IsWindows())
            return Windows.PathOfHandle(handle);
        return (OperatingSystem.IsLinux() ? Linux.PathOfHandle(handle) : null) ?? Unix.ResolveLinks(openedPath);
    }

    internal static string FinalDirectoryPath(string directory)
    {
        if (!OperatingSystem.IsWindows())
            return Unix.ResolveLinks(directory);

        using var handle = Windows.OpenDirectory(directory);
        return Windows.PathOfHandle(handle);
    }

    /// <summary>
    /// On Windows, whether a handle is a disk file rather than a pipe, a console or another device.
    /// Elsewhere it is always true: a device or FIFO there would already have been opened, and a FIFO's
    /// open blocks, so the policy relies on the roots holding no such file.
    /// </summary>
    internal static bool IsWindowsDiskFile(SafeFileHandle handle)
        => !OperatingSystem.IsWindows() || Windows.GetFileType(handle) == Windows.FileTypeDisk;

    /// <summary>
    /// Expands the 8.3 short names in the part of <paramref name="path"/> that exists. GetLongPathName
    /// fails for a path that does not exist, so the missing tail is put back after it.
    /// </summary>
    internal static string ExpandShortNames(string path)
    {
        if (!OperatingSystem.IsWindows())
            return path;

        var existing = path;
        var tail = "";
        while (!File.Exists(existing) && !Directory.Exists(existing))
        {
            var parent = Path.GetDirectoryName(existing);
            if (parent is null)
                return path;
            tail = Path.Combine(Path.GetFileName(existing), tail);
            existing = parent;
        }

        var expanded = Windows.LongPathName(existing) ?? existing;
        return tail.Length == 0 ? expanded : Path.Combine(expanded, tail).TrimEnd(Path.DirectorySeparatorChar);
    }

    [SupportedOSPlatform("windows")]
    private static partial class Windows
    {
        internal const uint FileTypeDisk = 1;
        private const uint FileReadAttributes = 0x80;
        private const uint ShareAll = 0x1 | 0x2 | 0x4;
        private const uint OpenExisting = 3;
        private const uint FlagBackupSemantics = 0x02000000;

        internal static SafeFileHandle OpenDirectory(string path)
        {
            var handle = CreateFileW(path, FileReadAttributes, ShareAll, IntPtr.Zero, OpenExisting,
                FlagBackupSemantics, IntPtr.Zero);
            if (handle.IsInvalid)
                throw new IOException($"Cannot open '{path}' (Win32 error {Marshal.GetLastPInvokeError()}).");
            return handle;
        }

        internal static string PathOfHandle(SafeFileHandle handle)
        {
            var buffer = new char[1024];
            while (true)
            {
                var length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Length, 0);
                if (length == 0)
                    throw new IOException($"Cannot read the final path of a handle (Win32 error {Marshal.GetLastPInvokeError()}).");
                if (length < buffer.Length)
                    return StripDevicePrefix(new string(buffer, 0, (int)length));
                buffer = new char[length + 1];
            }
        }

        internal static string? LongPathName(string path)
        {
            var buffer = new char[1024];
            var length = GetLongPathNameW(path, buffer, (uint)buffer.Length);
            if (length == 0 || length >= buffer.Length)
                return null;
            return new string(buffer, 0, (int)length);
        }

        private static string StripDevicePrefix(string path)
        {
            if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
                return @"\\" + path[8..];
            return path.StartsWith(@"\\?\", StringComparison.Ordinal) ? path[4..] : path;
        }

        [LibraryImport("kernel32.dll", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
        private static partial SafeFileHandle CreateFileW(string fileName, uint desiredAccess, uint shareMode,
            IntPtr securityAttributes, uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

        [LibraryImport("kernel32.dll", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
        private static partial uint GetFinalPathNameByHandleW(SafeFileHandle file, [Out] char[] path, uint length,
            uint flags);

        [LibraryImport("kernel32.dll", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
        private static partial uint GetLongPathNameW(string shortPath, [Out] char[] longPath, uint length);

        [LibraryImport("kernel32.dll")]
        internal static partial uint GetFileType(SafeFileHandle file);
    }

    [SupportedOSPlatform("linux")]
    private static partial class Linux
    {
        /// <summary>
        /// The target of the handle's /proc/self/fd link, or null when there is none. The handle is
        /// duplicated to learn a descriptor number, so the marshaller holds a reference on it for the
        /// call and the number never outlives the duplicate.
        /// </summary>
        internal static string? PathOfHandle(SafeFileHandle handle)
        {
            var descriptor = Dup(handle);
            if (descriptor < 0)
                throw new IOException($"Cannot duplicate a file handle (errno {Marshal.GetLastPInvokeError()}).");
            try
            {
                var procLink = $"/proc/self/fd/{descriptor}";
                return File.Exists(procLink) || Directory.Exists(procLink) ? new FileInfo(procLink).LinkTarget : null;
            }
            finally
            {
                Close(descriptor);
            }
        }

        [LibraryImport("libc", EntryPoint = "dup", SetLastError = true)]
        private static partial int Dup(SafeFileHandle descriptor);

        [LibraryImport("libc", EntryPoint = "close")]
        private static partial int Close(int descriptor);
    }

    private static class Unix
    {
        private const int MaxLinkHops = 40;

        /// <summary>
        /// Every link in the path followed, one component at a time. Unlike realpath(3), '..' is
        /// collapsed in the text first, as the policy's text gate does, so both read a path alike.
        /// </summary>
        internal static string ResolveLinks(string path)
        {
            var current = Path.GetFullPath(path);
            for (var hops = 0; hops < MaxLinkHops; hops++)
            {
                var root = Path.GetPathRoot(current) ?? "/";
                var segments = current[root.Length..].Split('/', StringSplitOptions.RemoveEmptyEntries);
                var prefix = root;
                string? next = null;
                for (var i = 0; i < segments.Length && next is null; i++)
                {
                    prefix = Path.Combine(prefix, segments[i]);
                    if (new FileInfo(prefix).LinkTarget is not { } target)
                        continue;
                    var absolute = Path.IsPathRooted(target) ? target : Path.Combine(Path.GetDirectoryName(prefix)!, target);
                    next = Path.GetFullPath(Path.Combine([absolute, .. segments[(i + 1)..]]));
                }

                if (next is null)
                    return current;
                current = next;
            }

            throw new IOException($"'{path}' goes through more than {MaxLinkHops} links.");
        }
    }
}
