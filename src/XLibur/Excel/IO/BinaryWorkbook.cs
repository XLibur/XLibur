using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;

namespace XLibur.Excel.IO;

internal static class BinaryWorkbook
{
    private const string WorkbookContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";

    internal static bool IsBinaryWorkbook(Package package)
    {
        return package.GetParts().Any(part =>
            string.Equals(part.ContentType, WorkbookContentType, StringComparison.OrdinalIgnoreCase));
    }

    internal static bool IsBinaryWorkbook(string path)
    {
        using var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var package = ReadOnlyPackageOpener.TryOpen(file);
        return package is not null && IsBinaryWorkbook(package);
    }

    internal static NotSupportedException NotSupported()
    {
        return new NotSupportedException(
            "The Excel Binary Workbook format (.xlsb) is not supported. Use an Excel Workbook (.xlsx) instead; " +
            "Excel can save an .xlsb file as .xlsx.");
    }
}
