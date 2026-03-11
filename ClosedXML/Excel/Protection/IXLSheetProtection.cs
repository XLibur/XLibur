#nullable disable


using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel;

public interface IXLSheetProtection : IXLElementProtection<XLSheetProtectionElements>
{
    IXLSheetProtection Protect(XLSheetProtectionElements allowedElements);

    IXLSheetProtection Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements);

    IXLSheetProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm, XLSheetProtectionElements allowedElements = XLSheetProtectionElements.SelectEverything);
}
