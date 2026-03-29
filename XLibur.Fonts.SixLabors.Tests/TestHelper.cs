using System.IO;
using System.Reflection;

namespace XLibur.Fonts.SixLabors.Tests;

internal static class TestHelper
{
    private static readonly Assembly Assembly = typeof(TestHelper).Assembly;

    public static Stream GetStreamFromResource(string resourceName)
    {
        var fullName = $"XLibur.Fonts.SixLabors.Tests.Resource.{resourceName}";
        return Assembly.GetManifestResourceStream(fullName)
               ?? throw new FileNotFoundException($"Embedded resource '{fullName}' not found.");
    }
}
