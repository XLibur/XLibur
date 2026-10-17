// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
//
// Ported from dotnet/runtime v10.0.0,
// src/libraries/System.IO.Packaging/src/System/IO/Packaging/ZipPackagePart.cs and ZipStreamManager.cs,
// and changed for XLibur as the remarks on XLZipPackage describe. See THIRD-PARTY-NOTICES.md.

using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;

namespace XLibur.Excel.IO.Packaging;

/// <summary>
/// A part of an <see cref="XLZipPackage"/>, stored in one zip entry. The first time the part is
/// opened to be written, its entry is replaced by one at the level of the save.
/// </summary>
internal sealed class XLZipPackagePart : PackagePart
{
    private readonly XLZipPackage _package;

    private ZipArchiveEntry? _entry;

    /// <summary>The level the part is written at.</summary>
    private readonly CompressionLevel _level;

    /// <summary>Was the entry created by this package, at <see cref="_level"/>?</summary>
    private bool _isAtLevel;

    /// <summary>Was the entry, still as the file had it, opened to be read?</summary>
    private bool _wasRead;

    internal XLZipPackagePart(XLZipPackage package, ZipArchiveEntry entry, Uri partUri, string contentType,
        CompressionLevel level, bool isAtLevel)
        : base(package, partUri, contentType, XLZipPackage.ToCompressionOption(level))
    {
        _package = package;
        _entry = entry;
        _level = level;
        _isAtLevel = isAtLevel;
    }

    protected override Stream? GetStreamCore(FileMode mode, FileAccess access)
    {
        if (_entry is null)
            return null;

        var write = mode == FileMode.Create || access != FileAccess.Read;
        if (!_isAtLevel && write)
        {
            // A part replaced to be created afresh has nothing to keep.
            _entry = _package.ReplaceEntry(_entry, _level, keepContent: mode != FileMode.Create);
            _isAtLevel = true;
            _wasRead = false;
        }
        else if (mode == FileMode.Create)
        {
            using var stream = _entry.Open();
            stream.SetLength(0);
        }
        else if (!_isAtLevel)
        {
            // Read where it is. A part that is later written is replaced then, without its content
            // read twice; one that is only read is stored at the level when the package closes.
            _wasRead = true;
        }

        var canRead = access != FileAccess.Write;
        var canWrite = access != FileAccess.Read;
        return new ZipWrappingStream(_entry.Open(), canRead, canWrite);
    }

    /// <summary>Stores the part at the level of the save if it was only read.</summary>
    internal void StoreAtLevelIfRead()
    {
        if (_entry is null || !_wasRead)
            return;

        _entry = _package.ReplaceEntry(_entry, _level, keepContent: true);
        _isAtLevel = true;
        _wasRead = false;
    }

    internal void DeleteEntry()
    {
        _entry?.Delete();
        _entry = null;
    }

    /// <summary>
    /// The stream of a zip entry, limited to the access the part was opened with.
    /// </summary>
    private sealed class ZipWrappingStream(Stream baseStream, bool canRead, bool canWrite) : Stream
    {
        public override bool CanRead => canRead && baseStream.CanRead;

        public override bool CanWrite => canWrite && baseStream.CanWrite;

        public override bool CanSeek => baseStream.CanSeek;

        // An entry opened in update mode is held uncompressed in memory, so its stream has a length.
        public override long Length => baseStream.Length;

        public override long Position
        {
            get => baseStream.Position;
            set => baseStream.Position = value;
        }

        public override int Read(byte[] buffer, int offset, int count) => baseStream.Read(buffer, offset, count);

        public override int Read(Span<byte> buffer) => baseStream.Read(buffer);

        public override void Write(byte[] buffer, int offset, int count) => baseStream.Write(buffer, offset, count);

        public override void Write(ReadOnlySpan<byte> buffer) => baseStream.Write(buffer);

        public override void SetLength(long value) => baseStream.SetLength(value);

        public override long Seek(long offset, SeekOrigin origin) => baseStream.Seek(offset, origin);

        public override void Flush() => baseStream.Flush();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                baseStream.Dispose();

            base.Dispose(disposing);
        }
    }
}
