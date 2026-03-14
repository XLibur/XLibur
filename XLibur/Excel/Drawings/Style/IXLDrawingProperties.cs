#nullable disable

namespace XLibur.Excel;

public interface IXLDrawingProperties
{
    XLDrawingAnchor Positioning { get; set; }
    IXLDrawingStyle SetPositioning(XLDrawingAnchor value);

}
