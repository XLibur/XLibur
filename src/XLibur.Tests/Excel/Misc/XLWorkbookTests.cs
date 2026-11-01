using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;
// ReSharper disable once InconsistentNaming
public class XLWorkbookTests
{
    [Test]
    public async Task Cell1()
    {
        var wb = new XLWorkbook();
        var cell = wb.Cell("ABC");
        await Assert.That(cell).IsNull();
    }

    [Test]
    public async Task Cell2()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        var cell = wb.Cell("Sheet1!Result");
        await Assert.That(cell).IsNotNull();
        await Assert.That(cell!.Value).IsEqualTo(1);
    }

    [Test]
    public async Task Cell3()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        var cell = wb.Cell("Sheet1!Result");
        await Assert.That(cell).IsNotNull();
        await Assert.That(cell!.Value).IsEqualTo(1);
    }

    [Test]
    public async Task Cells1()
    {
        var wb = new XLWorkbook();
        var cells = wb.Cells("ABC");
        await Assert.That(cells).IsNotNull();
        await Assert.That(cells.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task Cells2()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        var cells = wb.Cells("Sheet1!Result, ABC");
        await Assert.That(cells).IsNotNull();
        await Assert.That(cells.Count()).IsEqualTo(1);
        await Assert.That(cells.First().Value).IsEqualTo(1);
    }

    [Test]
    public async Task Cells3()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        var cells = wb.Cells("Sheet1!Result, ABC");
        await Assert.That(cells).IsNotNull();
        await Assert.That(cells.Count()).IsEqualTo(1);
        await Assert.That(cells.First().Value).IsEqualTo(1);
    }

    [Test]
    public async Task GetCellFromFullAddress()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var ws2 = wb.AddWorksheet("O'Sheet 2");
        var c1 = ws.Cell("C123");
        var c2 = ws2.Cell("B7");

        var c1Full = wb.Cell("Sheet1!C123");
        var c2Full = wb.Cell("'O'Sheet 2'!B7");

        await Assert.That(c1Full).IsEqualTo(c1);
        await Assert.That(c2Full).IsEqualTo(c2);
        await Assert.That(c1Full).IsNotNull();
        await Assert.That(c2Full).IsNotNull();
    }

    [Test]
    [Arguments("Sheet1")]
    [Arguments("Sheet1!")]
    [Arguments("Sheet2!")]
    [Arguments("Sheet2!C1")]
    [Arguments("Sheet1!ZZZ1")]
    [Arguments("Sheet1!A")]
    public async Task GetCellFromNonExistingFullAddress(string address)
    {
        var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");

        var c = wb.Cell(address);

        await Assert.That(c).IsNull();
    }

    [Test]
    public async Task GetRangeFromFullAddress()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var r1 = ws.Range("C123:D125");

        var r2 = wb.Range("Sheet1!C123:D125");

        await Assert.That(r2).IsSameReferenceAs(r1);
        await Assert.That(r2).IsNotNull();
    }

    [Test]
    [Arguments("Sheet2!C1:D2")]
    [Arguments("Sheet1!A")]
    public async Task GetRangeFromNonExistingFullAddress(string rangeAddress)
    {
        var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");

        var r = wb.Range(rangeAddress);

        await Assert.That(r).IsNull();
    }

    [Test]
    public async Task GetRangesFromFullAddress()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var r1 = ws.Ranges("A1:B2,C1:E3");

        var r2 = wb.Ranges("Sheet1!A1:B2,Sheet1!C1:E3");

        await Assert.That(r2.Count).IsEqualTo(2);
        await Assert.That(r2.First()).IsSameReferenceAs(r1.First());
        await Assert.That(r2.Last()).IsSameReferenceAs(r1.Last());
    }

    [Test]
    [Arguments("Sheet2!C1:D2,Sheet2!F1:G4")]
    [Arguments("Sheet1!A,Sheet1!B")]
    public async Task GetRangesFromNonExistingFullAddress(string rangesAddress)
    {
        var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");

        var r = wb.Ranges(rangesAddress);

        await Assert.That(r).IsNotNull();
        await Assert.That(r.Count).IsEqualTo(0);
    }

    [Test]
    public async Task Non_existent_defined_name_returns_null()
    {
        var wb = new XLWorkbook();
        var definedName = wb.DefinedName("ABC");
        await Assert.That(definedName).IsNull();
    }

    [Test]
    public async Task Sheet_specified_defined_name_is_retrieved_from_sheet_if_defined_there()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        var definedName = wb.DefinedName("Sheet1!Result");
        await Assert.That(definedName).IsNotNull();
        await Assert.That(definedName!.Ranges.Count).IsEqualTo(1);
        await Assert.That(definedName.Ranges.Cells().Count()).IsEqualTo(1);
        await Assert.That(definedName.Ranges.First().FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task Sheet_specified_defined_name_returns_null_if_not_defined_in_sheet_nor_workbook()
    {
        var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        var definedName = wb.DefinedName("Sheet1!Result");
        await Assert.That(definedName).IsNull();
    }

    [Test]
    public async Task Sheet_specified_defined_name_falls_back_to_workbook_scoped_defined_name_if_not_defined_in_sheet()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        var definedName = wb.DefinedName("Sheet1!Result");
        await Assert.That(definedName).IsNotNull();
        await Assert.That(definedName!.Ranges.Count).IsEqualTo(1);
        await Assert.That(definedName.Ranges.Cells().Count()).IsEqualTo(1);
        await Assert.That(definedName.Ranges.First().FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task Range1()
    {
        var wb = new XLWorkbook();
        var range = wb.Range("ABC");
        await Assert.That(range).IsNull();
    }

    [Test]
    public async Task Range2()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        var range = wb.Range("Sheet1!Result");
        await Assert.That(range).IsNotNull();
        await Assert.That(range!.Cells().Count()).IsEqualTo(1);
        await Assert.That(range.FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task Range3()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        var range = wb.Range("Sheet1!Result");
        await Assert.That(range).IsNotNull();
        await Assert.That(range!.Cells().Count()).IsEqualTo(1);
        await Assert.That(range.FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task Ranges1()
    {
        var wb = new XLWorkbook();
        var ranges = wb.Ranges("ABC");
        await Assert.That(ranges).IsNotNull();
        await Assert.That(ranges.Count).IsEqualTo(0);
    }

    [Test]
    public async Task Ranges2()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
        var ranges = wb.Ranges("Sheet1!Result, ABC");
        await Assert.That(ranges).IsNotNull();
        await Assert.That(ranges.Cells().Count()).IsEqualTo(1);
        await Assert.That(ranges.First().FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task Ranges3()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(1).AddToNamed("Result");
        var ranges = wb.Ranges("Sheet1!Result, ABC");
        await Assert.That(ranges).IsNotNull();
        await Assert.That(ranges.Cells().Count()).IsEqualTo(1);
        await Assert.That(ranges.First().FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task WbNamedCell()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
        await Assert.That(wb.Cell("TestCell")!.GetText()).IsEqualTo("Test");
        await Assert.That(ws.Cell("TestCell").GetText()).IsEqualTo("Test");
    }

    [Test]
    public async Task WbNamedCells()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
        ws.Cell(2, 1).SetValue("B").AddToNamed("Test2");
        var wbCells = wb.Cells("TestCell, Test2");
        await Assert.That(wbCells.First().GetText()).IsEqualTo("Test");
        await Assert.That(wbCells.Last().GetText()).IsEqualTo("B");

        var wsCells = ws.Cells("TestCell, Test2");
        await Assert.That(wsCells.First().GetText()).IsEqualTo("Test");
        await Assert.That(wsCells.Last().GetText()).IsEqualTo("B");
    }

    [Test]
    public async Task WbNamedRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("A");
        ws.Cell(2, 1).SetValue("B");
        var original = ws.Range("A1:A2");
        original.AddToNamed("TestRange");
        await Assert.That(wb.Range("TestRange")!.RangeAddress.ToString()).IsEqualTo(original.RangeAddress.ToStringFixed());
        await Assert.That(ws.Range("TestRange").RangeAddress.ToString()).IsEqualTo(original.RangeAddress.ToStringFixed());
    }

    [Test]
    public async Task WbNamedRanges()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).SetValue("A");
        ws.Cell(2, 1).SetValue("B");
        ws.Cell(3, 1).SetValue("C").AddToNamed("Test2");
        var original = ws.Range("A1:A2");
        original.AddToNamed("TestRange");
        var wbRanges = wb.Ranges("TestRange, Test2");
        await Assert.That(wbRanges.First().RangeAddress.ToString()).IsEqualTo(original.RangeAddress.ToStringFixed());
        await Assert.That(wbRanges.Last().RangeAddress.ToStringFixed()).IsEqualTo("$A$3:$A$3");

        var wsRanges = wb.Ranges("TestRange, Test2");
        await Assert.That(wsRanges.First().RangeAddress.ToString()).IsEqualTo(original.RangeAddress.ToStringFixed());
        await Assert.That(wsRanges.Last().RangeAddress.ToStringFixed()).IsEqualTo("$A$3:$A$3");
    }

    [Test]
    public async Task WbNamedRangesOneString()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        wb.DefinedNames.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");

        var wbRanges = ws.Ranges("TestRange");
        await Assert.That(wbRanges.First().RangeAddress.ToStringFixed()).IsEqualTo("$A$1:$A$1");
        await Assert.That(wbRanges.Last().RangeAddress.ToStringFixed()).IsEqualTo("$A$3:$A$3");

        var wsRanges = ws.Ranges("TestRange");
        await Assert.That(wsRanges.First().RangeAddress.ToStringFixed()).IsEqualTo("$A$1:$A$1");
        await Assert.That(wsRanges.Last().RangeAddress.ToStringFixed()).IsEqualTo("$A$3:$A$3");
    }

    [Test]
    public async Task WbProtect1()
    {
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        wb.Protect();
        await Assert.That(wb.LockStructure).IsTrue();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsFalse();
    }

    [Test]
    public async Task WbProtect2()
    {
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        wb.Protect(XLWorkbookProtectionElements.Windows);
        await Assert.That(wb.LockStructure).IsTrue();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsFalse();
    }

    [Test]
    public async Task WbProtect3()
    {
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        wb.Protect("Abc@123");
        await Assert.That(wb.LockStructure).IsTrue();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsTrue();
        await Assert.That(() => wb.Protect()).Throws<InvalidOperationException>();
        await Assert.That(() => wb.Unprotect()).Throws<InvalidOperationException>();
        await Assert.That(() => wb.Unprotect("Cde@345")).Throws<ArgumentException>();
    }

    [Test]
    public async Task WbProtect4()
    {
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        wb.Protect();
        await Assert.That(wb.LockStructure).IsTrue();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsFalse();
        wb.Unprotect();
        wb.Protect("Abc@123");
        await Assert.That(wb.LockStructure).IsTrue();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsTrue();
    }

    [Test]
    public async Task WbProtect5()
    {
        using var wb = new XLWorkbook();
        wb.Worksheets.Add("Sheet1");
        wb.Protect("Abc@123", XLProtectionAlgorithm.DefaultProtectionAlgorithm, XLWorkbookProtectionElements.Windows);
        await Assert.That(wb.LockStructure).IsTrue();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsTrue();
        wb.Unprotect("Abc@123");
        await Assert.That(wb.LockStructure).IsFalse();
        await Assert.That(wb.LockWindows).IsFalse();
        await Assert.That(wb.IsPasswordProtected).IsFalse();
    }

    [Test]
    public async Task FileSharingProperties()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = "Hello world!";
            wb.FileSharing.ReadOnlyRecommended = true;
            wb.FileSharing.UserName = Environment.UserName;
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.FileSharing.ReadOnlyRecommended).IsTrue();
            await Assert.That(wb.FileSharing.UserName).IsEqualTo(Environment.UserName);
        }
    }
}
