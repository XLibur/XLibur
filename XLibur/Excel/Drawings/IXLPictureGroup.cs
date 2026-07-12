using System.Collections.Generic;
using System.IO;

namespace XLibur.Excel.Drawings;

/// <summary>
/// A group of pictures within a worksheet drawing (an <c>xdr:grpSp</c> group shape). Obtained from
/// <see cref="IXLPicture.Group"/>, <see cref="IXLWorksheet.PictureGroups"/>, or
/// <see cref="IXLPictures.Group"/>.
/// </summary>
public interface IXLPictureGroup
{
    /// <summary>The worksheet that owns this group.</summary>
    IXLWorksheet Worksheet { get; }

    /// <summary>The pictures that are direct members of this group.</summary>
    IEnumerable<IXLPicture> Pictures { get; }

    /// <summary>
    /// Add a picture to this group. The returned picture's <c>Width</c>/<c>Height</c>/<c>Left</c>/
    /// <c>Top</c> are interpreted in sheet-space and can be set before saving. The group must already
    /// exist in the saved drawing.
    /// </summary>
    IXLPicture Add(Stream stream, string name);

    /// <summary>Add a picture to this group, auto-naming it.</summary>
    IXLPicture Add(Stream stream);

    /// <summary>
    /// Remove a picture from this group. The group and its other members are preserved. Equivalent to
    /// calling <see cref="IXLPicture.Delete"/> on a grouped picture.
    /// </summary>
    void Remove(IXLPicture picture);
}
