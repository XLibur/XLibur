#nullable disable

namespace XLibur.Excel;

public interface IXLDrawingStyle
{
    //IXLDrawingFont Font { get; }
    IXLDrawingAlignment Alignment { get; }
    IXLDrawingColorsAndLines ColorsAndLines { get; }
    IXLDrawingSize Size { get; }
    IXLDrawingProtection Protection { get; }
    IXLDrawingProperties Properties { get; }
    IXLDrawingMargins Margins { get; }
    IXLDrawingWeb Web { get; }
}
