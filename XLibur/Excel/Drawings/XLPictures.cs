using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace XLibur.Excel.Drawings;

internal sealed class XLPictures : IXLPictures, IEnumerable<XLPicture>
{
    private readonly List<XLPicture> _pictures = [];
    private readonly XLWorksheet _worksheet;

    public XLPictures(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
        Deleted = new HashSet<string>();
        DeletedFromGroups = new List<(int Id, string? RelId)>();
    }

    public int Count
    {
        [DebuggerStepThrough]
        get => _pictures.Count;
    }

    /// <summary>Rel ids of deleted top-level pictures, whose whole anchor is removed on save.</summary>
    internal ICollection<string> Deleted { get; private set; }

    /// <summary>
    /// Deleted pictures that lived inside a group shape. Keyed by drawing id (to locate the exact
    /// <c>xdr:pic</c>) plus rel id (to drop the image part if unreferenced). Removed in place on save
    /// so the group and its other shapes survive.
    /// </summary>
    internal ICollection<(int Id, string? RelId)> DeletedFromGroups { get; }

    public IXLPicture Add(Stream stream)
    {
        var picture = new XLPicture(_worksheet, stream);
        _pictures.Add(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(Stream stream, string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        var picture = Add(stream);
        picture.Name = name;
        return picture;
    }

    public IXLPicture Add(Stream stream, XLPictureFormat format)
    {
        var picture = new XLPicture(_worksheet, stream, format);
        _pictures.Add(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(Stream stream, XLPictureFormat format, string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        var picture = Add(stream, format);
        picture.Name = name;
        return picture;
    }

    public IXLPicture Add(string imageFile)
    {
        ArgumentException.ThrowIfNullOrEmpty(imageFile);
        using var fs = File.OpenRead(imageFile);
        var picture = new XLPicture(_worksheet, fs);
        _pictures.Add(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(string imageFile, string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(imageFile);
        ArgumentException.ThrowIfNullOrEmpty(name);
        var picture = Add(imageFile);
        picture.Name = name;
        return picture;
    }

    public bool Contains(string pictureName)
    {
        ArgumentNullException.ThrowIfNull(pictureName);
        return _pictures.Any(p => string.Equals(p.Name, pictureName, StringComparison.OrdinalIgnoreCase));
    }

    public void Delete(IXLPicture picture)
    {
        Delete(picture.Name);
    }

    public void Delete(string pictureName)
    {
        ArgumentException.ThrowIfNullOrEmpty(pictureName);
        var picturesToDelete = _pictures
            .Where(picture => picture.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (picturesToDelete.Count == 0)
            throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");

        foreach (var picture in picturesToDelete)
        {
            if (picture.IsInGroup)
                DeletedFromGroups.Add((picture.Id, picture.RelId));
            else if (!string.IsNullOrEmpty(picture.RelId))
                Deleted.Add(picture.RelId);

            _pictures.Remove(picture);
        }
    }

    IEnumerator<IXLPicture> IEnumerable<IXLPicture>.GetEnumerator()
    {
        return _pictures.Cast<IXLPicture>().GetEnumerator();
    }

    public IEnumerator<XLPicture> GetEnumerator()
    {
        return ((IEnumerable<XLPicture>)_pictures).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXLPicture Picture(string pictureName)
    {
        ArgumentException.ThrowIfNullOrEmpty(pictureName);
        if (TryGetPicture(pictureName, out IXLPicture? p))
            return p!;

        throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");
    }

    public bool TryGetPicture(string pictureName, out IXLPicture? picture)
    {
        ArgumentNullException.ThrowIfNull(pictureName);
        var match = _pictures.FirstOrDefault(p => p.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase));
        if (match is not null)
        {
            picture = match;
            return true;
        }

        picture = null;
        return false;
    }

    internal IXLPicture Add(Stream stream, string name, int id)
    {
        var picture = (XLPicture)Add(stream);
        picture.SetName(name);
        picture.Id = id;
        return picture;
    }

    private string GetNextPictureName()
    {
        var pictureNumber = Count;
        while (_pictures.Any(p => p.Name == $"Picture {pictureNumber}"))
        {
            pictureNumber++;
        }

        return $"Picture {pictureNumber}";
    }
}
