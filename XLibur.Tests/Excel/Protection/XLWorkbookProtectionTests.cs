using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Extensions;
using static XLibur.Excel.XLProtectionAlgorithm;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Protection;

public class XLWorkbookProtectionTests
{
    public static IEnumerable<Algorithm> AllAlgorithms() => Enum.GetValues<Algorithm>();

    [Test]
    public async Task CanChangeProtectionAlgorithm()
    {
        using var ms = new MemoryStream();
        using (var stream = GetProtectedWorkbookStreamWithPassword())
        using (var wb = new XLWorkbook(stream))
        {
            await Assert.That(wb.Protection.Algorithm).IsEqualTo(Algorithm.SHA512);
            wb.Unprotect("12345");
            wb.Protect("12345");

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.IsPasswordProtected).IsTrue();
            await Assert.That(wb.Protection.Algorithm).IsEqualTo(Algorithm.SimpleHash);
        }
    }

    [Test]
    public async Task CanChangeToPasswordProtected()
    {
        using var ms = new MemoryStream();
        using (var stream = GetProtectedWorkbookStreamWithoutPassword())
        using (var wb = new XLWorkbook(stream))

        {
            wb.Unprotect();
            wb.Protection.Protect("12345");

            await Assert.That(wb.Protection.IsPasswordProtected).IsTrue();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Protection.IsPasswordProtected).IsTrue();
            await Assert.That(wb.Protection.Algorithm).IsEqualTo(Algorithm.SimpleHash);
            await Assert.That(wb.Protection.PasswordHash).IsNotEqualTo("");
        }
    }

    [Test]
    public async Task CanChangeToProtectedWithoutPassword()
    {
        using var ms = new MemoryStream();
        using (var stream = GetProtectedWorkbookStreamWithPassword())
        using (var wb = new XLWorkbook(stream))

        {
            wb.Unprotect("12345");
            wb.Protection.Protect();

            await Assert.That(wb.Protection.IsPasswordProtected).IsFalse();
            await Assert.That(wb.Protection.IsProtected).IsTrue();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Protection.IsPasswordProtected).IsFalse();
            await Assert.That(wb.Protection.IsProtected).IsTrue();
            await Assert.That(wb.Protection.Algorithm).IsEqualTo(Algorithm.SimpleHash);
            await Assert.That(wb.Protection.PasswordHash).IsEqualTo("");
        }
    }

    [Test]
    public async Task CannotUnprotectIfNoPassword()
    {
        using var stream = GetProtectedWorkbookStreamWithoutPassword();
        using var wb = new XLWorkbook(stream);
        var ex = await Assert.That(() => wb.Unprotect("dummy password")).Throws<ArgumentException>();
        await Assert.That(ex!.Message).IsEqualTo("Invalid password");
    }

    [Test]
    public async Task CannotUnprotectWithoutPassword()
    {
        using var stream = GetProtectedWorkbookStreamWithPassword();
        using var wb = new XLWorkbook(stream);
        var ex = await Assert.That(() => wb.Unprotect()).Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).IsEqualTo("The workbook structure is password protected");
    }

    // NUnit's [Theory] fed this from the Algorithm enum automatically; TUnit needs the
    // values supplied explicitly.
    [Test]
    [MethodDataSource(nameof(AllAlgorithms))]
    public async Task CanProtectWithPassword(Algorithm algorithm)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet();

            await Assert.That(wb.Protection.IsProtected).IsFalse();

            wb.Protection.Protect("12345", algorithm);

            wb.Protection.AllowNone();
            await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)).IsFalse();
            await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)).IsFalse();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Protection.IsPasswordProtected).IsTrue();
            await Assert.That(wb.Protection.IsProtected).IsTrue();

            await Assert.That(wb.Protection.Algorithm).IsEqualTo(algorithm);
            await Assert.That(wb.Protection.PasswordHash).IsNotEqualTo("");

            await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)).IsFalse();
            await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)).IsFalse();

            var ex = await Assert.That(() => wb.Unprotect("dummy password")).Throws<ArgumentException>();
            await Assert.That(ex!.Message).IsEqualTo("Invalid password");

            wb.Protection.Unprotect("12345");

            wb.Save();
        }
    }

    [Test]
    public async Task CanUnprotectWithoutPassword()
    {
        using var ms = new MemoryStream();
        using (var stream = GetProtectedWorkbookStreamWithoutPassword())
        using (var wb = new XLWorkbook(stream))
        {
            // Unprotect without password
            wb.Unprotect();

            await Assert.That(wb.Protection.IsProtected).IsFalse();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Protection.IsProtected).IsFalse();
        }
    }

    [Test]
    public async Task CanUnprotectWithPassword()
    {
        using var ms = new MemoryStream();
        using (var stream = GetProtectedWorkbookStreamWithPassword())
        using (var wb = new XLWorkbook(stream))
        {
            // Unprotect with password
            wb.Unprotect("12345");

            await Assert.That(wb.Protection.IsProtected).IsFalse();

            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Protection.IsProtected).IsFalse();
        }
    }

    [Test]
    public async Task CopyProtectionFromAnotherWorkbook()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\WorkbookProtection.xlsx"));
        using var wb1 = new XLWorkbook(stream);
        using var wb2 = new XLWorkbook();
        wb2.AddWorksheet();

        var p1 = wb1.Protection.CastTo<XLWorkbookProtection>();
        await Assert.That(p1.IsProtected).IsTrue();

        await Assert.That(wb2.Protection.IsProtected).IsFalse();
        var p2 = wb2.Protection.CopyFrom(wb1.Protection).CastTo<XLWorkbookProtection>();

        await Assert.That(p2.IsProtected).IsTrue();
        await Assert.That(p2.IsPasswordProtected).IsTrue();
        await Assert.That(p2.Algorithm).IsEqualTo(p1.Algorithm);
        await Assert.That(p2.PasswordHash).IsEqualTo(p1.PasswordHash);
        await Assert.That(p2.Base64EncodedSalt).IsEqualTo(p1.Base64EncodedSalt);
        await Assert.That(p2.SpinCount).IsEqualTo(p1.SpinCount);

        await Assert.That(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)).IsTrue();
        await Assert.That(p2.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)).IsFalse();

        await Assert.That(() => wb2.Unprotect()).Throws<InvalidOperationException>();
        wb2.Unprotect("Abc@123");
    }

    [Test]
    public async Task IXLProtectableTests()
    {
        using var wb = new XLWorkbook();
        Enumerable.Range(1, 5).ForEach(_ => wb.AddWorksheet());

        var list = new List<IXLProtectable> { wb };
        list.AddRange(wb.Worksheets);

        // The assertions were List.ForEach(el => Assert...) under NUnit. TUnit assertions
        // must be awaited, and List.ForEach only accepts an Action -- an async lambda there
        // would be async void and the assertion would be swallowed. Hence explicit loops.
        list.ForEach(el => el.Protect());

        foreach (var el in list)
        {
            await Assert.That(el.IsProtected).IsTrue();
            await Assert.That(el.IsPasswordProtected).IsFalse();
        }

        list.ForEach(el => el.Unprotect());

        foreach (var el in list)
        {
            await Assert.That(el.IsProtected).IsFalse();
            await Assert.That(el.IsPasswordProtected).IsFalse();
        }

        list.ForEach(el => el.Protect("password"));

        foreach (var el in list)
        {
            await Assert.That(el.IsProtected).IsTrue();
            await Assert.That(el.IsPasswordProtected).IsTrue();
        }

        list.ForEach(el => el.Unprotect("password"));

        foreach (var el in list)
        {
            await Assert.That(el.IsProtected).IsFalse();
            await Assert.That(el.IsPasswordProtected).IsFalse();
        }
    }

    [Test]
    public async Task LoadProtectionWithoutPasswordFromFile()
    {
        using var stream = GetProtectedWorkbookStreamWithoutPassword();
        using var wb = new XLWorkbook(stream);
        await Assert.That(wb.Protection.IsPasswordProtected).IsFalse();
        await Assert.That(wb.Protection.IsProtected).IsTrue();
        await Assert.That(wb.Protection.PasswordHash).IsEqualTo("");
        await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)).IsTrue();
        await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)).IsFalse();
    }

    [Test]
    public async Task LoadProtectionWithPasswordFromFile()
    {
        using var stream = GetProtectedWorkbookStreamWithPassword();
        using var wb = new XLWorkbook(stream);
        await Assert.That(wb.Protection.IsPasswordProtected).IsTrue();
        await Assert.That(wb.Protection.PasswordHash).IsNotEqualTo("");
        await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows)).IsTrue();
        await Assert.That(wb.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure)).IsFalse();
    }

    [Test]
    public async Task SetWorkbookProtectionCloning()
    {
        var wb1 = new XLWorkbook();
        var wb2 = new XLWorkbook();

        wb1.AddWorksheet();
        wb2.AddWorksheet();

        wb1.Protect("123", Algorithm.SHA512)
            .AllowElement(XLWorkbookProtectionElements.Windows)
            .DisallowElement(XLWorkbookProtectionElements.Structure);

        await Assert.That(wb1.Protection.IsProtected).IsTrue();

        await Assert.That(wb1.Protection.AllowedElements).IsEqualTo(XLWorkbookProtectionElements.Windows);

        wb2.Protection = wb1.Protection;

        await Assert.That(ReferenceEquals(wb1.Protection, wb2.Protection)).IsFalse();
        await Assert.That(wb2.Protection.IsProtected).IsTrue();
        await Assert.That(wb2.Protection.AllowedElements).IsEqualTo(XLWorkbookProtectionElements.Windows);
        await Assert.That(wb2.Protection.PasswordHash).IsEqualTo(wb1.Protection.PasswordHash);
    }

    private static Stream GetProtectedWorkbookStreamWithoutPassword() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithoutpassword.xlsx"));

    private static Stream GetProtectedWorkbookStreamWithPassword() => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Protection\protectstructurewithpassword.xlsx"));
}
