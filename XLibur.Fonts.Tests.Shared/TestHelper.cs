using System.IO;
using System.Reflection;

namespace XLibur.Fonts.Tests;

/// <summary>
/// Reads the test fonts each font engine test project embeds under <c>Resource\Fonts</c>.
/// </summary>
/// <remarks>
/// Linked into every font engine test project, so the resource prefix is the compiling assembly's
/// name rather than a namespace written here.
/// </remarks>
internal static class TestHelper
{
    private static readonly Assembly Assembly = typeof(TestHelper).Assembly;

    public static Stream GetStreamFromResource(string resourceName)
    {
        var fullName = $"{Assembly.GetName().Name}.Resource.{resourceName}";
        return Assembly.GetManifestResourceStream(fullName)
               ?? throw new FileNotFoundException($"Embedded resource '{fullName}' not found.");
    }
}
