using XLibur.Examples.Styles;
using System.IO;

namespace XLibur.Examples;

public static class StyleExamples
{
    public static void Create()
    {
        var path = Program.BaseCreatedDirectory;
        new StyleFont().Create(Path.Combine(path, "styleFont.xlsx"));
        new StyleFill().Create(Path.Combine(path, "styleFill.xlsx"));
        new StyleBorder().Create(Path.Combine(path, "styleBorder.xlsx"));
        new StyleAlignment().Create(Path.Combine(path, "styleAlignment.xlsx"));
        new StyleNumberFormat().Create(Path.Combine(path, "styleNumberFormat.xlsx"));
        new StyleIncludeQuotePrefix().Create(Path.Combine(path, "styleIncludeQuotePrefix.xlsx"));
    }
}
