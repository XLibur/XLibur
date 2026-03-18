using System;
using System.IO;
using System.IO.Packaging;

namespace XLibur.Tests.Utils;

internal static class ExcelDocsComparer
{
    public static bool Compare(string left, string right, out string message)
    {
        using var leftStream = File.OpenRead(left);
        using var rightStream = File.OpenRead(right);
        return Compare(leftStream, rightStream, out message);
    }

    public static bool Compare(Stream left, Stream right, out string message)
    {
        using var leftPackage = Package.Open(left, FileMode.Open, FileAccess.Read);
        using var rightPackage = Package.Open(right, FileMode.Open, FileAccess.Read);
        return PackageHelper.Compare(leftPackage, rightPackage, false, ExcludeMethod, out message);
    }

    private static bool ExcludeMethod(Uri uri)
    {
        // Exclude service data
        return uri.OriginalString.EndsWith(".rels") ||
               uri.OriginalString.EndsWith(".psmdcp");
    }
}
