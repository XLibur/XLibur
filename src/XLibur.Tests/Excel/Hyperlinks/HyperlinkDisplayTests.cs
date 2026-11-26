using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Hyperlinks;

/// <summary>
/// A hyperlink's <c>display</c> attribute survives a load and save (#963).
/// </summary>
public class HyperlinkDisplayTests
{
    [Test]
    public async Task External_hyperlink_keeps_its_loaded_display_text()
    {
        using var created = Save(ws => ws.Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/")));
        created.RewriteSheet1(xml => xml.Replace("ref=\"B7\"", "ref=\"B7\" display=\"Example site\"", StringComparison.Ordinal));

        using var saved = LoadAndSave(created);

        await Assert.That(saved.Sheet1Xml()).Contains("display=\"Example site\"");
    }

    [Test]
    public async Task External_hyperlink_drops_its_loaded_display_text_when_the_address_changes()
    {
        using var created = Save(ws => ws.Cell("B7").SetHyperlink(new XLHyperlink("http://example.com/")));
        created.RewriteSheet1(xml => xml.Replace("ref=\"B7\"", "ref=\"B7\" display=\"Example site\"", StringComparison.Ordinal));

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(created.ToArray())))
        {
            wb.Worksheet("Data").Cell("B7").GetHyperlink().ExternalAddress = new Uri("http://example.org/");
            wb.SaveAs(saved);
        }

        await Assert.That(saved.Sheet1Xml()).DoesNotContain("display=");
    }

    private static MemoryStream Save(Action<IXLWorksheet> build)
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            build(wb.AddWorksheet("Data"));
            wb.SaveAs(output);
        }

        output.Position = 0;
        return output;
    }

    private static MemoryStream LoadAndSave(MemoryStream package)
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(package.ToArray())))
            wb.SaveAs(output);

        output.Position = 0;
        return output;
    }
}
