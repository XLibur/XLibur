using System;
using System.IO;
using System.IO.Packaging;
using System.Xml;

namespace XLibur.Excel.IO;

/// <summary>
/// Opens a package for reading without the copy the OpenXML SDK makes of it.
/// </summary>
/// <remarks>
/// Opened read-only from a stream or a file, the SDK wraps the source so it cannot be written,
/// then copies the entire package to a temporary file on disk and reopens it from there. The copy
/// exists for one purpose: an external relationship whose target is not a valid URI has to be
/// rewritten before <see cref="System.IO.Packaging"/> will read the part it belongs to, and the SDK
/// can only rewrite a package it can write. On a small workbook that copy was most of the cost of
/// opening the package, and it wrote the caller's whole file to disk to read it.
/// <para>
/// Almost no package needs the rewrite. So the package is opened here directly, and its
/// relationship parts are checked with the rule the SDK applies. When one would be rewritten, or
/// anything at all goes wrong, this gives up and the caller opens through the SDK as before, which
/// then produces the same document or the same exception it always did.
/// </para>
/// </remarks>
internal static class ReadOnlyPackageOpener
{
    private const string TargetModeAttribute = "TargetMode";
    private const string TargetAttribute = "Target";

    /// <summary>
    /// Opens <paramref name="stream"/> as a read-only package, or returns null when the SDK has to
    /// open it instead. The stream is left open either way.
    /// </summary>
    internal static Package? TryOpen(Stream stream)
    {
        Package? package = null;
        try
        {
            package = Package.Open(stream, FileMode.Open, FileAccess.Read);
            if (!NeedsUriRewriting(package))
                return package;
        }
        catch (Exception e) when (e is FormatException or IOException or InvalidDataException or ArgumentException
                                      or XmlException or InvalidOperationException or NotSupportedException)
        {
            // Whatever the SDK makes of this package — including the exception it throws for a
            // damaged one — is the answer, so it gets to open the package itself.
        }

        package?.Close();
        return null;
    }

    /// <summary>
    /// Whether any relationship part holds an external target the SDK would rewrite: one that is
    /// empty or does not parse as a relative or absolute URI.
    /// </summary>
    /// <remarks>
    /// Mirrors <c>PackageUriHandlingExtensions.Update</c> in the SDK, including its lenient reading
    /// of <c>TargetMode</c>. Every relationship part is checked, where the SDK checks those of the
    /// parts it visits, so this can only fall back more often than needed, never less.
    /// </remarks>
    private static bool NeedsUriRewriting(Package package)
    {
        foreach (var part in package.GetParts())
        {
            if (PackUriHelper.IsRelationshipPartUri(part.Uri) && HasUnparsableExternalTarget(part))
                return true;
        }

        return false;
    }

    private static bool HasUnparsableExternalTarget(PackagePart relationshipPart)
    {
        using var stream = relationshipPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = PartXmlReader.Create(stream);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;

            if (!Enum.TryParse<TargetMode>(reader.GetAttribute(TargetModeAttribute), out var mode)
                || mode != TargetMode.External)
            {
                continue;
            }

            var target = reader.GetAttribute(TargetAttribute) ?? string.Empty;
            if (target.Length == 0 || !Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out _))
                return true;
        }

        return false;
    }
}
