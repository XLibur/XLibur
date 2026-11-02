using Microsoft.Win32.SafeHandles;

namespace XLibur.Operations.Paths;

/// <summary>
/// Decides which files an operation may read and write. Restricted, it confines every path to a set
/// of root directories; unrestricted, it only canonicalises (the CLI in the user's own shell).
/// </summary>
/// <remarks>
/// <para>
/// A restricted path passes three gates. The text gate refuses what cannot name an ordinary local
/// file: device and UNC paths, alternate data streams, trailing dots or spaces, reserved device names,
/// an extension outside <see cref="AllowedExtensions"/>, and a name not under a root, all before the
/// file system is touched. The link gate walks the path one
/// component at a time and follows each symbolic link or junction, so a link to a network share or out
/// of the roots is refused before anything outside is opened. The handle gate opens the file and asks
/// the operating system where the handle really is, and that answer decides containment. On Windows and
/// Linux it sees past 8.3 short names, case and any link swapped in after the walk. Elsewhere there is
/// no portable query of a handle's path, so this gate resolves the name again, and only the walk holds.
/// </para>
/// <para>
/// A write goes to a temporary file in the destination's directory, whose handle passes the same
/// gate, and is renamed over the destination when it is complete. A reader never sees half a file,
/// and a rename replaces a link at the destination rather than writing through it.
/// </para>
/// </remarks>
public sealed class PathPolicy
{
    public static IReadOnlyList<string> AllowedExtensions { get; } = [".xlsx", ".xlsm", ".csv", ".tsv", ".json"];

    /// <summary>
    /// How a name the caller wrote is compared with a root: case-insensitive except on Linux. It only
    /// refuses early; a name that passes is decided by <see cref="FinalComparison"/>.
    /// </summary>
    private static readonly StringComparison NameComparison =
        OperatingSystem.IsLinux() ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    /// <summary>
    /// How a path the operating system gave for a handle is compared with a canonical root. Both carry
    /// the names as stored on disk, so the compare is exact. Ignoring case would let a sibling such as
    /// <c>data</c> pass for the root <c>Data</c> in a case-sensitive directory, or <c>data\u00B5</c> for
    /// <c>data\u039C</c>, which .NET folds together and NTFS keeps apart. Off Windows and Linux the final path
    /// is a resolved name, not a handle's, so case is ignored there as the default file system does.
    /// </summary>
    private static readonly StringComparison FinalComparison =
        OperatingSystem.IsWindows() || OperatingSystem.IsLinux() ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    private const int MaxLinkHops = 32;
    private const int MaxPathLength = 1024;
    private const int MaxComponents = 64;

    private static readonly HashSet<string> ReservedNames = new(
        new[] { "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$", "COM\u00B9", "COM\u00B2", "COM\u00B3", "LPT\u00B9", "LPT\u00B2", "LPT\u00B3" }
            .Concat(Enumerable.Range(0, 10).SelectMany(i => new[] { $"COM{i}", $"LPT{i}" })),
        StringComparer.OrdinalIgnoreCase);

    private readonly string[]? _roots;

    /// <summary>Each root as given, made absolute, and as canonicalised, for the early refusal of names.</summary>
    private readonly string[] _nameRoots;

    private PathPolicy(string[]? roots, string[] nameRoots, bool allowWrite)
    {
        _roots = roots;
        _nameRoots = nameRoots;
        AllowWrite = allowWrite;
    }

    /// <summary>The canonical root directories, or null when the policy is unrestricted.</summary>
    public IReadOnlyList<string>? Roots => _roots;

    public bool AllowWrite { get; }

    public static PathPolicy Unrestricted(bool allowWrite = true) => new(null, [], allowWrite);

    /// <exception cref="DirectoryNotFoundException">A root is not an existing directory.</exception>
    public static PathPolicy Restricted(IEnumerable<string> roots, bool allowWrite)
    {
        var given = roots.Select(r => Path.GetFullPath(r)).ToArray();
        if (given.Length == 0)
            throw new ArgumentException("A restricted policy needs at least one root.", nameof(roots));
        var canonical = given.Select(CanonicalRoot).Distinct(StringComparer.FromComparison(FinalComparison)).ToArray();
        return new PathPolicy(canonical, given.Concat(canonical).ToArray(), allowWrite);
    }

    public ReadableFile OpenRead(string path)
    {
        var full = CheckText(path);
        if (_roots is not null)
            full = ResolveLinksInsideRoots(full);

        SafeFileHandle handle;
        try
        {
            handle = File.OpenHandle(full, FileMode.Open, FileAccess.Read, FileShare.Read);
        }
        catch (Exception e) when (e is FileNotFoundException or DirectoryNotFoundException)
        {
            throw OperationException.BadInput("file_not_found", $"No file at '{path}'.");
        }
        catch (UnauthorizedAccessException)
        {
            throw OperationException.BadInput("access_denied", $"'{path}' cannot be opened for reading. Is it a directory?");
        }

        try
        {
            if (!NativePaths.IsWindowsDiskFile(handle))
                throw OperationException.Forbidden("not_a_file", $"'{path}' is not a file on disk.");

            var final = NativePaths.FinalPath(handle, full);
            if (_roots is not null && !IsInsideCanonicalRoots(final))
                throw OutsideRoots(path);
            CheckExtension(final, path);
            return new ReadableFile(final, handle);
        }
        catch
        {
            handle.Dispose();
            throw;
        }
    }

    /// <summary>Checks that <paramref name="path"/> may be written and returns where the write will land.</summary>
    public WriteTarget ResolveWrite(string path, bool overwrite)
    {
        if (!AllowWrite)
            throw OperationException.Forbidden("writes_disabled", "Writes are off. Start the server with --allow-write to enable them.");

        var full = CheckText(path);
        var name = Path.GetFileName(full);
        var directory = Path.GetDirectoryName(full);
        if (name.Length == 0 || directory is null)
            throw OperationException.BadInput("path_invalid", $"'{path}' does not name a file.");
        if (_roots is not null)
            directory = ResolveLinksInsideRoots(directory);
        if (!Directory.Exists(directory))
            throw OperationException.BadInput("directory_not_found", $"The directory of '{path}' does not exist.");

        if (_roots is not null)
        {
            directory = NativePaths.FinalDirectoryPath(directory);
            if (!IsInsideCanonicalRoots(directory))
                throw OutsideRoots(path);
        }

        var destination = Path.Combine(directory, name);
        if (Directory.Exists(destination))
            throw OperationException.BadInput("path_is_directory", $"'{path}' is a directory.");

        var existing = new FileInfo(destination);
        if (_roots is not null && LinkTargetOf(destination) is not null)
            throw OperationException.Forbidden("destination_is_link", $"'{path}' is a link; refusing to replace it.");
        if (existing.Exists && !overwrite)
            throw OperationException.BadInput("file_exists", $"'{path}' exists. Pass overwrite to replace it.");

        return new WriteTarget(destination, overwrite);
    }

    /// <summary>Writes through a temporary file in the same directory that is renamed over the target when complete.</summary>
    public void WriteAtomically(WriteTarget target, Action<Stream> write)
    {
        var (destination, overwrite) = (target.FullPath, target.Overwrite);
        var directory = Path.GetDirectoryName(destination)!;
        var temp = Path.Combine(directory, $".{Path.GetFileName(destination)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            {
                if (_roots is not null)
                {
                    var final = NativePaths.FinalPath(stream.SafeFileHandle, temp);
                    if (!IsInsideCanonicalRoots(final))
                        throw OutsideRoots(destination);
                }

                write(stream);
                stream.Flush(flushToDisk: true);
            }

            File.Move(temp, destination, overwrite);
        }
        catch (IOException) when (!overwrite && File.Exists(destination))
        {
            throw OperationException.BadInput("file_exists", $"'{destination}' exists. Pass overwrite to replace it.");
        }
        finally
        {
            if (File.Exists(temp))
                File.Delete(temp);
        }
    }

    /// <summary>Whether two canonical paths are the same: exactly on Windows and Linux, ignoring case elsewhere.</summary>
    public static bool SamePath(string a, string b) => string.Equals(a, b, FinalComparison);

    private static string CanonicalRoot(string full)
    {
        if (!Directory.Exists(full))
            throw new DirectoryNotFoundException($"Root '{full}' is not an existing directory.");
        return NativePaths.FinalDirectoryPath(full);
    }

    /// <remarks>
    /// Restricted, a path that is not under a root by name is refused here, before anything touches the
    /// file system: a mapped network drive, or any directory outside, is never probed, so a refusal
    /// costs no network round trip and says nothing about what exists outside.
    /// </remarks>
    private string CheckText(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw OperationException.BadInput("path_empty", "A path is required.");
        if (path.Contains('\0'))
            throw OperationException.Forbidden("path_invalid", "A path cannot hold a NUL character.");
        if (path.Length > MaxPathLength || path.Split('/', '\\').Length > MaxComponents)
            throw OperationException.BadInput("path_too_long", $"A path may have at most {MaxPathLength} characters and {MaxComponents} components.");
        if (_roots is not null && IsDeviceOrUnc(path))
            throw OperationException.Forbidden("device_or_unc_path", $"'{path}' is a device or network path; only local paths inside the roots are allowed.");

        CheckComponents(path);
        var full = Path.GetFullPath(path, _roots is null ? Directory.GetCurrentDirectory() : _nameRoots[0]);
        if (_roots is not null && IsDeviceOrUnc(full))
            throw OperationException.Forbidden("device_or_unc_path", $"'{path}' is a device or network path.");
        CheckComponents(full);
        CheckExtension(full, path);
        if (_roots is not null && !IsInsideRootsByName(full))
            throw OutsideRoots(path);
        return full;
    }

    private static void CheckExtension(string full, string path)
    {
        var extension = Path.GetExtension(full);
        if (!AllowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            throw OperationException.Forbidden("extension_not_allowed",
                $"'{path}' has extension '{extension}'; allowed: {string.Join(", ", AllowedExtensions)}.");
    }

    private static void CheckComponents(string path)
    {
        var body = path;
        if (body.StartsWith(@"\\?\", StringComparison.Ordinal) || body.StartsWith(@"\\.\", StringComparison.Ordinal))
            body = body[4..];
        var colon = body.IndexOf(':');
        var driveColon = colon == 1 && char.IsAsciiLetter(body[0]);
        if (colon >= 0 && (!driveColon || body.IndexOf(':', 2) >= 0))
            throw OperationException.Forbidden("alternate_data_stream", $"'{path}' holds a ':' outside a drive letter (an alternate data stream).");

        foreach (var segment in body.Split('/', '\\'))
        {
            if (segment.Length == 0 || segment is "." or "..")
                continue;
            if (segment[^1] is '.' or ' ')
                throw OperationException.Forbidden("trailing_dot_or_space", $"'{path}' has a component ending in a dot or a space, which Windows strips.");

            var stem = segment.Split('.')[0].TrimEnd(' ');
            if (ReservedNames.Contains(stem))
                throw OperationException.Forbidden("reserved_device_name", $"'{path}' uses the reserved device name '{stem}'.");
        }
    }

    private static bool IsDeviceOrUnc(string path)
        => path.Length >= 2 && IsSeparator(path[0]) && IsSeparator(path[1])
           || path.StartsWith(@"\??\", StringComparison.Ordinal);

    private static bool IsSeparator(char c) => c is '/' or '\\';

    private string ResolveLinksInsideRoots(string full)
    {
        var current = NativePaths.ExpandShortNames(full);
        for (var hops = 0; ; hops++)
        {
            var link = FirstLink(current);
            if (link is null)
                break;
            if (hops >= MaxLinkHops)
                throw OperationException.Forbidden("link_loop", $"'{full}' goes through more than {MaxLinkHops} links.");

            var (linkPath, target, rest) = link.Value;
            if (target.StartsWith(@"\??\", StringComparison.Ordinal) || target.StartsWith(@"\\?\", StringComparison.Ordinal))
                target = target[4..];
            if (IsDeviceOrUnc(target))
                throw OperationException.Forbidden("link_outside_roots", $"'{linkPath}' links to the network or device path '{target}'.");

            var absolute = Path.IsPathRooted(target) ? target : Path.Combine(Path.GetDirectoryName(linkPath)!, target);
            current = Path.GetFullPath(rest.Length == 0 ? absolute : Path.Combine(absolute, rest));
            if (!IsInsideRootsByName(current))
                throw OutsideRoots(full);
            current = NativePaths.ExpandShortNames(current);
        }

        if (!IsInsideRootsByName(current))
            throw OutsideRoots(full);
        return current;
    }

    private static (string LinkPath, string Target, string Remainder)? FirstLink(string path)
    {
        var root = Path.GetPathRoot(path) ?? "";
        var segments = path[root.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        var prefix = root;
        for (var i = 0; i < segments.Length; i++)
        {
            prefix = Path.Combine(prefix, segments[i]);
            var target = LinkTargetOf(prefix);
            if (target is not null)
                return (prefix, target, string.Join(Path.DirectorySeparatorChar, segments[(i + 1)..]));
        }

        return null;
    }

    private static string? LinkTargetOf(string path)
    {
        try
        {
            return new FileInfo(path).LinkTarget;
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
    }

    private bool IsInsideRootsByName(string path) => _nameRoots.Any(root => IsUnder(path, root, NameComparison));

    private bool IsInsideCanonicalRoots(string path) => _roots!.Any(root => IsUnder(path, root, FinalComparison));

    private static bool IsUnder(string path, string root, StringComparison comparison)
    {
        var trimmed = Path.TrimEndingDirectorySeparator(root);
        if (string.Equals(path, trimmed, comparison))
            return true;
        var withSeparator = Path.EndsInDirectorySeparator(trimmed) ? trimmed : trimmed + Path.DirectorySeparatorChar;
        return path.StartsWith(withSeparator, comparison);
    }

    private OperationException OutsideRoots(string path)
        => OperationException.Forbidden("path_outside_roots",
            $"'{path}' is outside the allowed roots ({string.Join(", ", _roots!)}).");
}

/// <summary>
/// A destination that passed the write gates. Only <see cref="PathPolicy.ResolveWrite"/> makes one,
/// so a write cannot be aimed at a path the policy has not checked.
/// </summary>
public sealed class WriteTarget
{
    internal WriteTarget(string fullPath, bool overwrite)
    {
        FullPath = fullPath;
        Overwrite = overwrite;
    }

    public string FullPath { get; }

    public bool Overwrite { get; }
}

/// <summary>A file that passed the read gates, held open so its content is what was checked.</summary>
public sealed class ReadableFile(string fullPath, SafeFileHandle handle) : IDisposable
{
    public string FullPath { get; } = fullPath;

    public byte[] ReadAllBytes(long maxBytes)
    {
        var length = RandomAccess.GetLength(handle);
        if (length > maxBytes)
            throw OperationException.BadInput("file_too_large",
                $"'{FullPath}' is {length:N0} bytes; the limit is {maxBytes:N0}.");

        var bytes = new byte[length];
        var read = 0;
        while (read < bytes.Length)
        {
            var n = RandomAccess.Read(handle, bytes.AsSpan(read), read);
            if (n == 0)
                break;
            read += n;
        }

        return read == bytes.Length ? bytes : bytes[..read];
    }

    public void Dispose() => handle.Dispose();
}
