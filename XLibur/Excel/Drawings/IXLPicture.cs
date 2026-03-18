using System;
using System.Drawing;
using System.IO;

namespace XLibur.Excel.Drawings;

public interface IXLPicture : IDisposable
{
    IXLCell BottomRightCell { get; }

    /// <summary>
    /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
    /// Default value is "jpeg"
    /// </summary>
    XLPictureFormat Format { get; }

    /// <summary>
    /// Current width of the picture in pixels.
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// Current height of the picture in pixels.
    /// </summary>
    int Height { get; set; }

    int Id { get; }

    MemoryStream ImageStream { get; }

    int Left { get; set; }

    string Name { get; set; }

    /// <summary>
    /// Original height of the picture in pixels.
    /// </summary>
    int OriginalHeight { get; }

    /// <summary>
    /// Original width of the picture in pixels.
    /// </summary>
    int OriginalWidth { get; }

    XLPicturePlacement Placement { get; set; }

    int Top { get; set; }

    IXLCell TopLeftCell { get; }

    IXLWorksheet Worksheet { get; }

    /// <summary>
    /// Create a copy of the picture on a different worksheet.
    /// </summary>
    /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
    /// <returns>A created copy of the picture.</returns>
    IXLPicture CopyTo(IXLWorksheet targetSheet);

    /// <summary>
    /// Deletes this picture.
    /// </summary>
    void Delete();

    /// <summary>
    /// Create a copy of the picture on the same worksheet.
    /// </summary>
    /// <returns>A created copy of the picture.</returns>
    IXLPicture Duplicate();

    Point GetOffset(XLMarkerPosition position);

    IXLPicture MoveTo(int left, int top);

    IXLPicture MoveTo(IXLCell cell);

    IXLPicture MoveTo(IXLCell cell, int xOffset, int yOffset);

    IXLPicture MoveTo(IXLCell cell, Point offset);

    IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell);

    IXLPicture MoveTo(IXLCell fromCell, int fromCellXOffset, int fromCellYOffset, IXLCell toCell, int toCellXOffset, int toCellYOffset);

    IXLPicture MoveTo(IXLCell fromCell, Point fromOffset, IXLCell toCell, Point toOffset);

    IXLPicture Scale(double factor, bool relativeToOriginal = false);

    IXLPicture ScaleHeight(double factor, bool relativeToOriginal = false);

    IXLPicture ScaleWidth(double factor, bool relativeToOriginal = false);

    IXLPicture WithPlacement(XLPicturePlacement value);

    IXLPicture WithSize(int width, int height);
}
