using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using static XLibur.Excel.XLProtectionAlgorithm;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Protection;

public class XLSheetProtectionTests
{
    [Test]
    public async Task AllowEverything()
    {
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowedElements = XLSheetProtectionElements.Everything;

            foreach (var element in Enum.GetValues<XLSheetProtectionElements>())
                await Assert.That(ws.Protection.AllowedElements.HasFlag(element)).IsTrue().Because(element.ToString());
        }

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowElement(XLSheetProtectionElements.Everything);

            foreach (var element in Enum.GetValues<XLSheetProtectionElements>())
                await Assert.That(ws.Protection.AllowedElements.HasFlag(element)).IsTrue().Because(element.ToString());
        }

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowEverything();

            foreach (var element in Enum.GetValues<XLSheetProtectionElements>())
                await Assert.That(ws.Protection.AllowedElements.HasFlag(element)).IsTrue().Because(element.ToString());
        }
    }

    [Test]
    public async Task AllowNothing()
    {
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowedElements = XLSheetProtectionElements.None;

            foreach (var element in Enum.GetValues<XLSheetProtectionElements>()
                         .Where(e => e != XLSheetProtectionElements.None))

                await Assert.That(ws.Protection.AllowedElements.HasFlag(element)).IsFalse().Because(element.ToString());
        }

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowNone();

            foreach (var element in Enum.GetValues<XLSheetProtectionElements>()
                         .Where(e => e != XLSheetProtectionElements.None))

                await Assert.That(ws.Protection.AllowedElements.HasFlag(element)).IsFalse().Because(element.ToString());
        }
    }

    [Test]
    public async Task ChangeHashingAlgorithm()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Protect("123");

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Protection.IsProtected).IsTrue();
            await Assert.That(ws.Protection.Algorithm).IsEqualTo(Algorithm.SimpleHash);

            ws.Unprotect("123");
            ws.Protect("123", Algorithm.SHA512);
            wb.Save();
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Protection.IsProtected).IsTrue();
            await Assert.That(ws.Protection.Algorithm).IsEqualTo(Algorithm.SHA512);

            await Assert.That(() => ws.Unprotect("123")).ThrowsNothing();
        }
    }

    [Test]
    public async Task CopyProtectionFromAnotherSheet()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\SheetProtection.xlsx"));
        using var wb = new XLWorkbook(stream);
#pragma warning disable S2068 // Test data — hard-coded password is not a real credential
        var ws1 = wb.Worksheet("Protected Password = 123");
#pragma warning restore S2068
        var p1 = ws1.Protection.CastTo<XLSheetProtection>();
        await Assert.That(p1.IsProtected).IsTrue();

        var ws2 = ws1.CopyTo("New worksheet");
        await Assert.That(ws2.Protection.IsProtected).IsFalse();
        var p2 = ws2.Protection.CopyFrom(p1).CastTo<XLSheetProtection>();

        await Assert.That(p2.IsProtected).IsTrue();
        await Assert.That(p2.IsPasswordProtected).IsTrue();
        await Assert.That(p2.Algorithm).IsEqualTo(p1.Algorithm);
        await Assert.That(p2.PasswordHash).IsEqualTo(p1.PasswordHash);
        await Assert.That(p2.Base64EncodedSalt).IsEqualTo(p1.Base64EncodedSalt);
        await Assert.That(p2.SpinCount).IsEqualTo(p1.SpinCount);

        await Assert.That(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertColumns)).IsTrue();
        await Assert.That(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertRows)).IsTrue();
        await Assert.That(p2.AllowedElements.HasFlag(XLSheetProtectionElements.InsertHyperlinks)).IsFalse();

        await Assert.That(() => ws2.Unprotect()).Throws<InvalidOperationException>();
        ws2.Unprotect("123");
    }

    [Test]
    public async Task SetWorksheetProtectionCloning()
    {
        var ws1 = new XLWorkbook().AddWorksheet();
        var ws2 = new XLWorkbook().AddWorksheet();

        ws1.Protect("123")
            .AllowElement(XLSheetProtectionElements.FormatEverything)
            .DisallowElement(XLSheetProtectionElements.FormatCells);

        await Assert.That(ws1.Protection.AllowedElements).IsEqualTo(XLSheetProtectionElements.FormatColumns | XLSheetProtectionElements.FormatRows | XLSheetProtectionElements.SelectEverything);

        ws2.Protection = ws1.Protection;

        await Assert.That(ReferenceEquals(ws1.Protection, ws2.Protection)).IsFalse();
        await Assert.That(ws2.Protection.IsProtected).IsTrue();
        await Assert.That(ws2.Protection.AllowedElements).IsEqualTo(XLSheetProtectionElements.FormatColumns | XLSheetProtectionElements.FormatRows | XLSheetProtectionElements.SelectEverything);
        await Assert.That(((XLSheetProtection)ws2.Protection).PasswordHash).IsEqualTo(((XLSheetProtection)ws1.Protection).PasswordHash);
    }

    [Test]
    public async Task TestUnprotectWorksheetWithNoPassword()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Sheet1");
        await Assert.That(ws.Protection.IsProtected).IsTrue();
        ws.Unprotect();
        await Assert.That(ws.Protection.IsProtected).IsFalse();
    }

    [Test]
    public async Task TestWorksheetWithSHA512Protection()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Sheet2");
        await Assert.That(ws.Protection.IsProtected).IsTrue();

        // Password required
        await Assert.That(() => ws.Unprotect()).Throws<InvalidOperationException>();

        await Assert.That(ws.Protection.Algorithm).IsEqualTo(Algorithm.SHA512);
        ws.Unprotect("abc");
        await Assert.That(ws.Protection.IsProtected).IsFalse();
    }
}
