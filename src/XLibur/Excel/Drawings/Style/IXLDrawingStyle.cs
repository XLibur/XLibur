namespace XLibur.Excel;

public interface IXLDrawingStyle
{
    IXLDrawingAlignment Alignment { get; }
    IXLDrawingColorsAndLines ColorsAndLines { get; }
    IXLDrawingSize Size { get; }
    IXLDrawingProtection Protection { get; }
    IXLDrawingProperties Properties { get; }
    IXLDrawingMargins Margins { get; }
    IXLDrawingWeb Web { get; }
}
