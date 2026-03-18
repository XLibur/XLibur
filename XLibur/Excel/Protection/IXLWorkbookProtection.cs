using static XLibur.Excel.XLProtectionAlgorithm;

namespace XLibur.Excel;

public interface IXLWorkbookProtection : IXLElementProtection<XLWorkbookProtectionElements>
{
    IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements);

    IXLWorkbookProtection Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements);

    IXLWorkbookProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm, XLWorkbookProtectionElements allowedElements = XLWorkbookProtectionElements.Windows);
}
