using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace XLibur.Excel.Drawings;

/// <summary>
/// A lightweight view over the pictures of a single group, identified by its stable
/// <see cref="XLPictureGroup.GroupKey"/>. Membership is read live from the worksheet's picture
/// collection, so the view stays correct as pictures are added to or removed from the group.
/// </summary>
internal sealed class XLPictureGroupView : IXLPictureGroup, IEquatable<XLPictureGroupView>
{
    private readonly XLWorksheet _worksheet;
    private readonly long _groupKey;

    internal XLPictureGroupView(XLWorksheet worksheet, long groupKey)
    {
        _worksheet = worksheet;
        _groupKey = groupKey;
    }

    public IXLWorksheet Worksheet => _worksheet;

    public IEnumerable<IXLPicture> Pictures =>
        ((IEnumerable<XLPicture>)(XLPictures)_worksheet.Pictures).Where(p => p.GroupInfo?.GroupKey == _groupKey);

    public IXLPicture Add(Stream stream, string name) => AddCore(stream, name);

    public IXLPicture Add(Stream stream) => AddCore(stream, null);

    public void Remove(IXLPicture picture)
    {
        ArgumentNullException.ThrowIfNull(picture);
        if (picture is not XLPicture member)
            throw new ArgumentException("The picture was not created by this library.", nameof(picture));
        if (member.GroupInfo?.GroupKey != _groupKey)
            throw new ArgumentException("The picture is not a member of this group.", nameof(picture));

        member.Delete();
    }

    public bool Equals(XLPictureGroupView? other) =>
        other is not null && ReferenceEquals(_worksheet, other._worksheet) && _groupKey == other._groupKey;

    public override bool Equals(object? obj) => Equals(obj as XLPictureGroupView);

    public override int GetHashCode() => HashCode.Combine(_worksheet, _groupKey);

    private IXLPicture AddCore(Stream stream, string? name)
    {
        var pictures = (XLPictures)_worksheet.Pictures;
        var member = ((IEnumerable<XLPicture>)pictures).FirstOrDefault(p => p.GroupInfo?.GroupKey == _groupKey)
                     ?? throw new InvalidOperationException("The group has no members to anchor the new picture to.");

        return pictures.AddToGroup(member, stream, name);
    }
}
