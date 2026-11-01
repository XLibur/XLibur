using System.IO;

namespace XLibur.Examples;

public static class ExampleHelper
{
    private static string GetTempFilePath()
    {
        return Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
    }

    public static string GetTempFilePath(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        var tempFilePath = GetTempFilePath();
        return Path.ChangeExtension(tempFilePath, extension);
    }
}
