using System.Collections.Generic;
using System.IO;

namespace XLibur.Excel.Drawings;

public interface IXLPictures : IEnumerable<IXLPicture>
{
    int Count { get; }

    IXLPicture Add(Stream stream);

    IXLPicture Add(Stream stream, string name);

    IXLPicture Add(Stream stream, XLPictureFormat format);

    IXLPicture Add(Stream stream, XLPictureFormat format, string name);

    IXLPicture Add(string imageFile);

    IXLPicture Add(string imageFile, string name);

    bool Contains(string pictureName);

    void Delete(string pictureName);

    void Delete(IXLPicture picture);

    IXLPicture Picture(string pictureName);

    bool TryGetPicture(string pictureName, out IXLPicture? picture);

    /// <summary>
    /// Group two or more free-floating pictures on this worksheet into a new group shape. The
    /// pictures keep their on-sheet positions and sizes. Pictures must already be saved and use
    /// free-floating placement (call <see cref="IXLPicture.MoveTo(int, int)"/> first if needed).
    /// </summary>
    IXLPictureGroup Group(params IXLPicture[] pictures);
}
