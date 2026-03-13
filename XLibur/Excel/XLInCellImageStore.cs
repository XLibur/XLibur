using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using XLibur.Excel.Drawings;

namespace XLibur.Excel;

/// <summary>
/// Workbook-level store for in-cell image blobs. Deduplicates by SHA256 hash.
/// </summary>
internal sealed class XLInCellImageStore
{
    private readonly List<(MemoryStream Stream, XLPictureFormat Format)> _images = new();
    private readonly Dictionary<string, int> _hashToIndex = new(StringComparer.Ordinal);

    /// <summary>
    /// Number of images in the store.
    /// </summary>
    internal int Count => _images.Count;

    /// <summary>
    /// Add an image blob to the store, deduplicating by content hash.
    /// </summary>
    /// <param name="imageStream">Stream containing image data. Position is read from current to end.</param>
    /// <param name="format">Image format.</param>
    /// <returns>0-based index of the image in the store.</returns>
    internal int Add(Stream imageStream, XLPictureFormat format)
    {
        var ms = new MemoryStream();
        imageStream.CopyTo(ms);
        ms.Position = 0;

        var hash = ComputeHash(ms);
        ms.Position = 0;

        if (_hashToIndex.TryGetValue(hash, out var existingIndex))
        {
            ms.Dispose();
            return existingIndex;
        }

        var index = _images.Count;
        _images.Add((ms, format));
        _hashToIndex[hash] = index;
        return index;
    }

    /// <summary>
    /// Add an image blob directly (used during load). No dedup check.
    /// </summary>
    internal int AddDirect(MemoryStream stream, XLPictureFormat format)
    {
        var index = _images.Count;
        _images.Add((stream, format));

        stream.Position = 0;
        var hash = ComputeHash(stream);
        stream.Position = 0;
        _hashToIndex[hash] = index;

        return index;
    }

    /// <summary>
    /// Get an image blob by index.
    /// </summary>
    internal (MemoryStream Stream, XLPictureFormat Format) GetImage(int index)
    {
        return _images[index];
    }

    private static string ComputeHash(MemoryStream ms)
    {
        using var sha = SHA256.Create();
        var hashBytes = sha.ComputeHash(ms);
        return Convert.ToBase64String(hashBytes);
    }
}
