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

    /// <summary>Whether this picture is a member of a drawing group (<c>xdr:grpSp</c>).</summary>
    bool IsInGroup { get; }

    /// <summary>The group this picture belongs to, or <c>null</c> if it is not grouped.</summary>
    IXLPictureGroup? Group { get; }

    // ── Styling ──────────────────────────────────────────────────────────
    //
    // Each property reads what the file has, or what was assigned since. A property that is never
    // assigned is never written, so styling XLibur does not model, and styling nobody changed,
    // stays as the file has it. Assigning null removes the value from the file.

    /// <summary>The colour of the border (<c>a:ln</c> fill), or <c>null</c> for none set.</summary>
    /// <remarks>Assigning a colour clears <see cref="NoBorder"/>.</remarks>
    XLColor? BorderColor { get; set; }

    /// <summary>The width of the border in points (<c>a:ln/@w</c>), or <c>null</c> for none set.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is negative or above 1584 points.</exception>
    double? BorderWidthPt { get; set; }

    /// <summary>The dash pattern of the border (<c>a:prstDash</c>), or <c>null</c> for none set.</summary>
    XLLineDashStyle? BorderDash { get; set; }

    /// <summary>
    /// Whether the picture explicitly has no border (<c>a:ln/a:noFill</c>). Setting it removes
    /// <see cref="BorderColor"/>.
    /// </summary>
    bool NoBorder { get; set; }

    /// <summary>
    /// The colour behind the image, which shows through its transparent parts (<c>spPr</c> solid
    /// fill), or <c>null</c> for none set. Assigning a colour clears <see cref="NoFill"/>.
    /// </summary>
    XLColor? FillColor { get; set; }

    /// <summary>
    /// Whether the picture explicitly has no fill (<c>a:noFill</c>). Setting it removes
    /// <see cref="FillColor"/>.
    /// </summary>
    bool NoFill { get; set; }

    /// <summary>
    /// The rotation in degrees clockwise, 0 up to 360 (<c>a:xfrm/@rot</c>), or <c>null</c> for none.
    /// The picture turns about its centre; <see cref="Width"/>, <see cref="Height"/>,
    /// <see cref="Left"/> and <see cref="Top"/> keep describing the unrotated box. Any angle can be
    /// assigned and is brought into 0 up to 360.
    /// </summary>
    double? Rotation { get; set; }

    /// <summary>Whether the picture is mirrored left to right (<c>a:xfrm/@flipH</c>).</summary>
    bool? FlipHorizontal { get; set; }

    /// <summary>Whether the picture is mirrored top to bottom (<c>a:xfrm/@flipV</c>).</summary>
    bool? FlipVertical { get; set; }

    /// <summary>
    /// How transparent the image is, 0 (opaque) to 100 percent (<c>a:alphaModFix</c>), or
    /// <c>null</c> for none set.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 100.</exception>
    double? Transparency { get; set; }

    /// <summary>
    /// The recolour preset, or <c>null</c> when the file has none. A recolour that is no preset reads
    /// <see cref="XLPictureRecolor.Custom"/>.
    /// </summary>
    /// <remarks>
    /// <see cref="XLPictureRecolor.Washout"/> is Excel's brightness 70 % and contrast −70 %, so
    /// assigning it sets <see cref="Brightness"/> and <see cref="Contrast"/> too, and assigning
    /// either of them afterwards makes the picture read as no recolour.
    /// </remarks>
    /// <exception cref="ArgumentException"><see cref="XLPictureRecolor.Custom"/> is assigned.</exception>
    XLPictureRecolor? Recolor { get; set; }

    /// <summary>
    /// The brightness, −100 to 100 percent, 0 being the image's own (<c>a:lum/@bright</c>), or
    /// <c>null</c> for none set.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside −100 to 100.</exception>
    double? Brightness { get; set; }

    /// <summary>
    /// The contrast, −100 to 100 percent, 0 being the image's own (<c>a:lum/@contrast</c>), or
    /// <c>null</c> for none set.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside −100 to 100.</exception>
    double? Contrast { get; set; }

    /// <summary>How much of each edge of the image is cut away (<c>a:srcRect</c>), or <c>null</c> for none.</summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// An edge is 100 percent or more, or the two edges of one side together cut all of it away.
    /// </exception>
    XLCrop? Crop { get; set; }

    /// <summary>
    /// The shadow outside the picture (<c>a:outerShdw</c>), or <c>null</c> for none. A file whose
    /// effects are anything but one outer shadow reads <c>null</c>, and keeps its effects until
    /// this is assigned.
    /// </summary>
    XLPictureShadow? Shadow { get; set; }

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
