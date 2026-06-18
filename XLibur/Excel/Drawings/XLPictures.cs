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
    private long _nextGroupKey = 1;

    public XLPictures(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
        Deleted = new HashSet<string>();
        DeletedFromGroups = new List<(int Id, string? RelId)>();
        PendingGroups = new List<XLPendingGroup>();
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

    /// <summary>Groups to create on the next save (members are moved into a new group shape).</summary>
    internal ICollection<XLPendingGroup> PendingGroups { get; }

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

    /// <summary>
    /// Add a picture as a new member of the same group as <paramref name="groupMember"/>. The new
    /// picture inherits the group's composed transform, so its <c>Width</c>/<c>Height</c>/<c>Left</c>/
    /// <c>Top</c> are interpreted in sheet-space; the writer creates the <c>xdr:pic</c> inside the
    /// group on save. (A first-class group API is added in a later phase.)
    /// </summary>
    internal XLPicture AddToGroup(XLPicture groupMember, Stream stream, string? name = null)
    {
        var source = groupMember.GroupInfo
                     ?? throw new ArgumentException("The picture is not part of a group.", nameof(groupMember));

        var picture = (XLPicture)(name is null ? Add(stream) : Add(stream, name));
        picture.Placement = XLPicturePlacement.FreeFloating;
        picture.GroupInfo = new XLPictureGroup
        {
            ScaleX = source.ScaleX,
            ScaleY = source.ScaleY,
            OffsetX = source.OffsetX,
            OffsetY = source.OffsetY,
            GroupId = source.GroupId,
            GroupKey = source.GroupKey,
            IsNew = true,
        };
        return picture;
    }

    /// <summary>Allocate a stable, per-worksheet identity for a group (see <see cref="XLPictureGroup.GroupKey"/>).</summary>
    internal long AllocateGroupKey() => _nextGroupKey++;

    /// <summary>
    /// Group existing free-floating pictures into a new group shape. The group's bounding box is the
    /// union of the members' sheet rectangles; the members keep their on-sheet positions and sizes
    /// (the group uses an identity child coordinate space). The group is built on the next save.
    /// (A first-class group API is added in a later phase.)
    /// </summary>
    public IXLPictureGroup Group(params IXLPicture[] pictures)
    {
        if (pictures is null || pictures.Length < 2)
            throw new ArgumentException("A group needs at least two pictures.", nameof(pictures));

        var members = new XLPicture[pictures.Length];
        for (var i = 0; i < pictures.Length; i++)
            members[i] = (XLPicture)pictures[i];

        var wb = _worksheet.Workbook;
        long minX = long.MaxValue, minY = long.MaxValue, maxX = long.MinValue, maxY = long.MinValue;

        foreach (var member in members)
        {
            if (!ReferenceEquals(member.Worksheet, _worksheet))
                throw new ArgumentException("All pictures must belong to this worksheet.", nameof(pictures));
            if (member.IsInGroup)
                throw new ArgumentException($"Picture '{member.Name}' is already in a group.", nameof(pictures));
            if (member.Placement != XLPicturePlacement.FreeFloating)
                throw new NotSupportedException(
                    $"Only free-floating pictures can be grouped; '{member.Name}' is '{member.Placement}'. Call MoveTo(left, top) first.");
            if (string.IsNullOrEmpty(member.RelId))
                throw new NotSupportedException($"Picture '{member.Name}' must be saved before it can be grouped.");

            var x = ToEmu(member.Left, wb.DpiX);
            var y = ToEmu(member.Top, wb.DpiY);
            minX = Math.Min(minX, x);
            minY = Math.Min(minY, y);
            maxX = Math.Max(maxX, x + ToEmu(member.Width, wb.DpiX));
            maxY = Math.Max(maxY, y + ToEmu(member.Height, wb.DpiY));
        }

        var groupKey = AllocateGroupKey();
        foreach (var member in members)
        {
            // Identity child space (chOff/chExt == off/ext), so a child's coordinates equal its
            // sheet coordinates: composed scale 1, composed offset 0.
            member.GroupInfo = new XLPictureGroup
            {
                ScaleX = 1.0,
                ScaleY = 1.0,
                OffsetX = 0.0,
                OffsetY = 0.0,
                GroupKey = groupKey,
                LoadedWidthPx = member.Width,
                LoadedHeightPx = member.Height,
                LoadedLeftPx = member.Left,
                LoadedTopPx = member.Top,
            };
        }

        PendingGroups.Add(new XLPendingGroup([.. members], minX, minY, maxX - minX, maxY - minY));
        return new XLPictureGroupView(_worksheet, groupKey);
    }

    private static long ToEmu(int pixels, double dpi) => Convert.ToInt64(914400L * pixels / dpi);

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

/// <summary>
/// A group to be created on the next save: its member pictures and the group's bounding box in EMU
/// (used for the group's <c>off</c>/<c>ext</c> and <c>chOff</c>/<c>chExt</c>).
/// </summary>
internal sealed record XLPendingGroup(
    IReadOnlyList<XLPicture> Members,
    long OffsetX,
    long OffsetY,
    long ExtentCx,
    long ExtentCy);
