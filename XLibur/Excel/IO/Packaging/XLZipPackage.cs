// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
//
// Ported from dotnet/runtime v10.0.0,
// src/libraries/System.IO.Packaging/src/System/IO/Packaging/ZipPackage.cs, and changed for XLibur
// as the remarks on XLZipPackage describe. See THIRD-PARTY-NOTICES.md.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Xml;

namespace XLibur.Excel.IO.Packaging;

/// <summary>
/// The zip package a save of a loaded workbook edits: <see cref="ZipPackage"/>, except that every
/// part the save opens is stored at the level the save asks for (#688).
/// </summary>
/// <remarks>
/// <see cref="ZipPackage"/> rewrites a part into its existing zip entry, and an entry's compression
/// is fixed when the entry is created, so a loaded part kept the level the file gave it whatever
/// <see cref="SaveOptions.CompressionLevel"/> said. This package deletes the entry the first time a
/// part is opened and creates it again at the requested level.
/// <para>
/// A part the save never opens is never touched: <see cref="ZipArchiveMode.Update"/> moves its
/// compressed bytes as they are. That is how a sheet the save keeps whole (#702) keeps its bytes.
/// Opening an entry in that mode, even only to read it, makes the archive write it again, so a part
/// that is opened is a part the save writes, and is stored at the requested level too. The package
/// relationships and the core properties are the exception: they are stored uncompressed, as
/// <see cref="ZipPackage"/> creates them in every package (<see cref="LevelFor"/>).
/// </para>
/// <para>
/// Changes from <see cref="ZipPackage"/>: only an existing package opened for reading and writing is
/// supported, which is all a save needs. Interleaved parts, which only XPS writes, are not read; their
/// pieces stay in the archive as unknown entries, as do entries that do not name a valid part. Content
/// types are compared by their text, with the type and subtype compared as <see cref="ZipPackage"/>
/// does. Part names are validated with the public <see cref="PackUriHelper"/>, which applies the same
/// rules.
/// </para>
/// </remarks>
internal sealed class XLZipPackage : Package
{
    private const string ContentTypesFile = "[Content_Types].xml";

    private const string ForwardSlash = "/";

    private const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";

    private static readonly Uri PackageRelationshipsPartUri = new("/_rels/.rels", UriKind.Relative);

    private readonly ZipArchive _zipArchive;

    private readonly CompressionLevel _level;

    private readonly ContentTypeHelper _contentTypeHelper;

    private readonly IgnoredItemHelper _ignoredItemHelper;

    /// <summary>Every part this package created an object for, to store the ones only read.</summary>
    private readonly List<XLZipPackagePart> _parts = [];

    /// <summary>
    /// Opens the package in <paramref name="stream"/> for reading and writing. Each part the save
    /// opens is stored at <paramref name="level"/>. The stream is left open when the package closes.
    /// </summary>
    internal XLZipPackage(Stream stream, CompressionLevel level)
        : base(FileAccess.ReadWrite)
    {
        if (stream.CanSeek && stream.Length == 0)
            throw new FileFormatException("A zero-size stream is not a valid zip archive.");

        _level = level;
        stream.Position = 0;
        try
        {
            _zipArchive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
        }
        catch (InvalidDataException)
        {
            throw new FileFormatException("The package contains corrupted data.");
        }

        try
        {
            _ignoredItemHelper = new IgnoredItemHelper(_zipArchive);
            _contentTypeHelper = new ContentTypeHelper(_zipArchive, level, _ignoredItemHelper);

            // As Package.Open does for a ZipPackage: every part is read now, which GetPartCore relies on.
            GetParts();
        }
        catch
        {
            _zipArchive.Dispose();
            throw;
        }
    }

    protected override PackagePart CreatePartCore(Uri partUri, string contentType, CompressionOption compressionOption)
    {
        ArgumentNullException.ThrowIfNull(contentType);

        var normalizedName = NormalizedName(partUri);

        // An ignored entry that names the same part would show up as a second part next time.
        _ignoredItemHelper.Delete(normalizedName);

        // The base class passes a validated part URI, whose original string is the part name.
        var zipItemName = partUri.OriginalString[1..];
        var level = ToCompressionLevel(compressionOption);
        var entry = _zipArchive.CreateEntry(zipItemName, level);

        _contentTypeHelper.AddContentType(partUri, normalizedName, contentType);

        return Track(new XLZipPackagePart(this, entry, partUri, contentType, level, isAtLevel: true));
    }

    protected override PackagePart? GetPartCore(Uri partUri)
    {
        // As in ZipPackage: every part is read in GetPartsCore, so a part not found then does not exist.
        return null;
    }

    protected override void DeletePartCore(Uri partUri)
    {
        var normalizedName = NormalizedName(partUri);

        // The base class still holds the part when it calls this.
        if (PartExists(partUri) && GetPart(partUri) is XLZipPackagePart part)
            part.DeleteEntry();

        _ignoredItemHelper.Delete(normalizedName);
        _contentTypeHelper.DeleteContentType(normalizedName);
    }

    protected override PackagePart[] GetPartsCore()
    {
        var parts = new List<PackagePart>(_zipArchive.Entries.Count);
        foreach (var entry in _zipArchive.Entries)
        {
            if (!IsZipItemValidOpcPart(entry.FullName))
                continue;

            var partUri = new Uri(ForwardSlash + entry.FullName, UriKind.Relative);
            if (!TryNormalizedName(partUri, out var normalizedName))
            {
                // Not a valid part name, so no part can ever map to this entry.
                continue;
            }

            var contentType = _contentTypeHelper.GetContentType(partUri, normalizedName);
            if (contentType is null)
            {
                // A part with no content type. It is kept out of the way, as a part added later under
                // the same name or with the same extension would make it show up again.
                _ignoredItemHelper.AddItemForAtomicPart(partUri, normalizedName, entry.FullName);
                continue;
            }

            parts.Add(Track(new XLZipPackagePart(this, entry, partUri, contentType,
                LevelFor(partUri, contentType), isAtLevel: false)));
        }

        return parts.ToArray();
    }

    protected override void FlushCore()
    {
        _contentTypeHelper.SaveToFile();
    }

    protected override void Dispose(bool disposing)
    {
        try
        {
            if (disposing)
            {
                // A part the save only read is written again by the archive all the same; it goes
                // in at the requested level like the rest. Every part stream is closed by now.
                foreach (var part in _parts)
                    part.StoreAtLevelIfRead();

                _contentTypeHelper.SaveToFile();
                _zipArchive.Dispose();
            }
        }
        finally
        {
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// Replaces <paramref name="entry"/> with an entry of the same name at <paramref name="level"/>,
    /// with the same content when <paramref name="keepContent"/>, or empty.
    /// </summary>
    internal ZipArchiveEntry ReplaceEntry(ZipArchiveEntry entry, CompressionLevel level, bool keepContent)
    {
        MemoryStream? content = null;
        if (keepContent)
        {
            // Not sized from entry.Length, which an entry already opened in update mode refuses.
            content = new MemoryStream();
            using var old = entry.Open();
            old.CopyTo(content);
        }

        var name = entry.FullName;
        var lastWriteTime = entry.LastWriteTime;

        // Deleted before the new entry is created, so the archive never holds the name twice.
        entry.Delete();
        var replacement = _zipArchive.CreateEntry(name, level);
        replacement.LastWriteTime = lastWriteTime;

        if (content is not null)
        {
            using var stream = replacement.Open();
            content.Position = 0;
            content.CopyTo(stream);
        }

        return replacement;
    }

    /// <summary>
    /// The level a loaded part is written at: the save's, except for the package relationships and
    /// the core properties.
    /// </summary>
    /// <remarks>
    /// <see cref="System.IO.Packaging"/> creates both of those uncompressed in every package, so a new
    /// workbook's save stores them whatever the level. Deflating parts of a few hundred bytes at
    /// <see cref="CompressionLevel.Optimal"/> cost more than it saved: on a workbook of 8,000 numbers
    /// whose properties were changed, the round trip took 8% longer.
    /// </remarks>
    private CompressionLevel LevelFor(Uri partUri, string contentType) =>
        PackUriHelper.ComparePartUri(partUri, PackageRelationshipsPartUri) == 0
        || string.Equals(contentType, CorePropertiesContentType, StringComparison.OrdinalIgnoreCase)
            ? CompressionLevel.NoCompression
            : _level;

    private XLZipPackagePart Track(XLZipPackagePart part)
    {
        _parts.Add(part);
        return part;
    }

    /// <summary>
    /// Is the zip item a part? Leaves out the content types item, and entries with a leading or a
    /// trailing slash, such as the folders some zip tools add.
    /// </summary>
    private static bool IsZipItemValidOpcPart(string zipItemName)
    {
        if (zipItemName.StartsWith(ContentTypesFile, StringComparison.OrdinalIgnoreCase))
            return false;

        return !zipItemName.StartsWith(ForwardSlash, StringComparison.Ordinal)
               && !zipItemName.EndsWith(ForwardSlash, StringComparison.Ordinal);
    }

    private static string NormalizedName(Uri partUri) =>
        PackUriHelper.GetNormalizedPartUri(partUri).OriginalString;

    private static bool TryNormalizedName(Uri partUri, out string normalizedName)
    {
        try
        {
            normalizedName = NormalizedName(partUri);
            return true;
        }
        catch (ArgumentException)
        {
            normalizedName = string.Empty;
            return false;
        }
    }

    /// <summary>The extension of a part name, without the dot, or empty.</summary>
    private static string ExtensionOf(string partName)
    {
        var slash = partName.LastIndexOf('/');
        var dot = partName.LastIndexOf('.');
        return dot > slash ? partName[(dot + 1)..] : string.Empty;
    }

    private static CompressionLevel ToCompressionLevel(CompressionOption compressionOption) => compressionOption switch
    {
        CompressionOption.NotCompressed => CompressionLevel.NoCompression,
        CompressionOption.Normal => CompressionLevel.Optimal,
        CompressionOption.Maximum => CompressionLevel.SmallestSize,
        CompressionOption.Fast or CompressionOption.SuperFast => CompressionLevel.Fastest,
        _ => throw new ArgumentOutOfRangeException(nameof(compressionOption), compressionOption, null)
    };

    internal static CompressionOption ToCompressionOption(CompressionLevel level) => level switch
    {
        CompressionLevel.NoCompression => CompressionOption.NotCompressed,
        CompressionLevel.Fastest => CompressionOption.Fast,
        CompressionLevel.SmallestSize => CompressionOption.Maximum,
        _ => CompressionOption.Normal
    };

    /// <summary>
    /// Keeps the content types of the package, <c>[Content_Types].xml</c>: a default content type per
    /// extension, and an override per part whose content type differs from its extension's default.
    /// </summary>
    private sealed class ContentTypeHelper
    {
        private const string TypesNamespaceUri = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string TypesTagName = "Types";
        private const string DefaultTagName = "Default";
        private const string ExtensionAttributeName = "Extension";
        private const string ContentTypeAttributeName = "ContentType";
        private const string OverrideTagName = "Override";
        private const string PartNameAttributeName = "PartName";

        private readonly ZipArchive _zipArchive;

        private readonly CompressionLevel _level;

        private readonly IgnoredItemHelper _ignoredItemHelper;

        /// <summary>Extensions as the file spelled them, compared ignoring case.</summary>
        private readonly Dictionary<string, string> _defaults = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>Overrides by normalized part name: the part name as given, and its content type.</summary>
        private readonly Dictionary<string, (string PartName, string ContentType)> _overrides = new(StringComparer.Ordinal);

        private ZipArchiveEntry? _contentTypesEntry;

        private bool _dirty;

        internal ContentTypeHelper(ZipArchive zipArchive, CompressionLevel level, IgnoredItemHelper ignoredItemHelper)
        {
            _zipArchive = zipArchive;
            _level = level;
            _ignoredItemHelper = ignoredItemHelper;

            _contentTypesEntry = FindContentTypesEntry(zipArchive);
            if (_contentTypesEntry is null)
                return;

            using (var stream = _contentTypesEntry.Open())
                Parse(stream);

            // Reading the entry made the archive write it again, so it is written at the save's level.
            _dirty = true;
        }

        /// <summary>
        /// Records the content type of a new part: as its extension's default when the extension has
        /// none yet, as nothing more when the default matches, or else as an override.
        /// </summary>
        internal void AddContentType(Uri partUri, string normalizedName, string contentType)
        {
            var extension = ExtensionOf(partUri.OriginalString);
            if (extension.Length == 0
                || (_defaults.TryGetValue(extension, out var defaultType) && !AreTypeAndSubTypeEqual(defaultType, contentType)))
            {
                _overrides[normalizedName] = (partUri.OriginalString, contentType);
                _dirty = true;
            }
            else if (defaultType is null)
            {
                _defaults.Add(extension, contentType);
                _dirty = true;

                // Ignored items with this extension would show up as parts the next time.
                _ignoredItemHelper.DeleteItemsWithExtension(extension);
            }
        }

        internal string? GetContentType(Uri partUri, string normalizedName)
        {
            if (_overrides.TryGetValue(normalizedName, out var entry))
                return entry.ContentType;

            return _defaults.GetValueOrDefault(ExtensionOf(partUri.OriginalString));
        }

        internal void DeleteContentType(string normalizedName)
        {
            if (_overrides.Remove(normalizedName))
                _dirty = true;
        }

        internal void SaveToFile()
        {
            if (!_dirty)
                return;

            // Created again rather than overwritten: an entry's level is fixed when it is created,
            // and a shorter text would not truncate the old one.
            _contentTypesEntry?.Delete();
            _contentTypesEntry = _zipArchive.CreateEntry(ContentTypesFile, _level);

            using (var stream = _contentTypesEntry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8 }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement(TypesTagName, TypesNamespaceUri);

                foreach (var (extension, contentType) in _defaults)
                {
                    writer.WriteStartElement(DefaultTagName);
                    writer.WriteAttributeString(ExtensionAttributeName, extension);
                    writer.WriteAttributeString(ContentTypeAttributeName, contentType);
                    writer.WriteEndElement();
                }

                foreach (var (partName, contentType) in _overrides.Values)
                {
                    writer.WriteStartElement(OverrideTagName);
                    writer.WriteAttributeString(PartNameAttributeName, partName);
                    writer.WriteAttributeString(ContentTypeAttributeName, contentType);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            _dirty = false;
        }

        private static ZipArchiveEntry? FindContentTypesEntry(ZipArchive zipArchive)
        {
            foreach (var entry in zipArchive.Entries)
            {
                if (string.Equals(entry.FullName, ContentTypesFile, StringComparison.OrdinalIgnoreCase))
                    return entry;
            }

            return null;
        }

        /// <summary>
        /// Reads the content types. The package was loaded before it is saved, through
        /// <see cref="ZipPackage"/>, which checks the file against its schema; this reads only what
        /// the save needs.
        /// </summary>
        private void Parse(Stream stream)
        {
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });
            reader.MoveToContent();
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != TypesTagName
                                                       || reader.NamespaceURI != TypesNamespaceUri)
                throw new XmlException("The content types of the package do not start with a Types element.");

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1 || reader.NamespaceURI != TypesNamespaceUri)
                    continue;

                var contentType = reader.GetAttribute(ContentTypeAttributeName);
                if (contentType is null)
                    continue;

                if (reader.LocalName == DefaultTagName)
                {
                    var extension = reader.GetAttribute(ExtensionAttributeName);
                    if (!string.IsNullOrEmpty(extension))
                        _defaults[extension] = contentType;
                }
                else if (reader.LocalName == OverrideTagName)
                {
                    var partName = reader.GetAttribute(PartNameAttributeName);
                    if (!string.IsNullOrEmpty(partName)
                        && TryNormalizedName(new Uri(partName, UriKind.Relative), out var normalizedName))
                        _overrides[normalizedName] = (partName, contentType);
                }
            }
        }

        /// <summary>Do two content types have the same type and subtype, ignoring case and parameters?</summary>
        private static bool AreTypeAndSubTypeEqual(string first, string second)
        {
            var firstMediaType = first.AsSpan();
            var secondMediaType = second.AsSpan();
            var firstEnd = firstMediaType.IndexOf(';');
            var secondEnd = secondMediaType.IndexOf(';');
            if (firstEnd >= 0) firstMediaType = firstMediaType[..firstEnd];
            if (secondEnd >= 0) secondMediaType = secondMediaType[..secondEnd];
            return firstMediaType.Trim().Equals(secondMediaType.Trim(), StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// Keeps the zip items that name a valid part but have no content type. They are deleted when a
    /// part of the same name is created, or a default content type is added for their extension, so
    /// that they do not show up as parts the next time the package is opened.
    /// </summary>
    private sealed class IgnoredItemHelper(ZipArchive zipArchive)
    {
        /// <summary>Zip item names by normalized part name.</summary>
        private readonly Dictionary<string, List<string>> _items = new(StringComparer.Ordinal);

        /// <summary>Normalized part names by extension, compared ignoring case.</summary>
        private readonly Dictionary<string, List<string>> _extensions = new(StringComparer.OrdinalIgnoreCase);

        internal void AddItemForAtomicPart(Uri partUri, string normalizedName, string zipItemName)
        {
            if (!_items.TryGetValue(normalizedName, out var zipItemNames))
                _items.Add(normalizedName, zipItemNames = []);

            zipItemNames.Add(zipItemName);

            var extension = ExtensionOf(partUri.OriginalString);
            if (!_extensions.TryGetValue(extension, out var normalizedNames))
                _extensions.Add(extension, normalizedNames = []);

            normalizedNames.Add(normalizedName);
        }

        internal void Delete(string normalizedName)
        {
            if (!_items.Remove(normalizedName, out var zipItemNames))
                return;

            foreach (var zipItemName in zipItemNames)
                zipArchive.GetEntry(zipItemName)?.Delete();
        }

        internal void DeleteItemsWithExtension(string extension)
        {
            if (!_extensions.Remove(extension, out var normalizedNames))
                return;

            foreach (var normalizedName in normalizedNames)
                Delete(normalizedName);
        }
    }
}
