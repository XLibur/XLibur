using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Excel;
using ClosedXML.Parser;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.NamedRanges;

public class NamedRangesTests
{
    [Test]
    public async Task Formula_must_be_valid()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet();
        await Assert.That(() => wb.DefinedNames.Add("Test", "SUM(Sheet7!A4")).Throws<ParsingException>();
    }

    [Test]
    public async Task CanEvaluateNamedMultiRange()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet1");
        ws1.Range("A1:C1").Value = 1;
        ws1.Range("A3:C3").Value = 3;
        wb.DefinedNames.Add("TEST", ws1.Ranges("A1:C1,A3:C3"));

        ws1.Cell(2, 1).FormulaA1 = "=SUM(TEST)";

        await Assert.That((double)ws1.Cell(2, 1).Value).IsEqualTo(12.0).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task CanGetNamedFromAnother()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").SetValue(1).AddToNamed("value1");

        await Assert.That(wb.Cell("value1")!.Value).IsEqualTo(1);
        await Assert.That(wb.Range("value1")!.FirstCell().Value).IsEqualTo(1);

        await Assert.That(ws1.Cell("value1").Value).IsEqualTo(1);
        await Assert.That(ws1.Range("value1").FirstCell().Value).IsEqualTo(1);

        var ws2 = wb.Worksheets.Add("Sheet2");

        ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

        await Assert.That(wb.Cell("value2")!.Value).IsEqualTo(1);
        await Assert.That(wb.Range("value2")!.FirstCell().Value).IsEqualTo(1);

        await Assert.That(ws2.Cell("value1").Value).IsEqualTo(1);
        await Assert.That(ws2.Range("value1").FirstCell().Value).IsEqualTo(1);

        await Assert.That(ws2.Cell("value2").Value).IsEqualTo(1);
        await Assert.That(ws2.Range("value2").FirstCell().Value).IsEqualTo(1);
    }

    [Test]
    public async Task CanGetValidNamedRanges()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet 1");
        var ws2 = wb.Worksheets.Add("Sheet 2");
        var ws3 = wb.Worksheets.Add("Sheet'3");

        ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
        ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
        ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
        ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
        wb.DefinedNames.Add("Named range 5", new XLRanges
        {
            ws1.Range("A5:D5"),
            ws3.Range("A5:D5")
        });

        ws2.Delete();
        ws3.Delete();

        var globalValidRanges = wb.DefinedNames.ValidNamedRanges().ToList();
        var globalInvalidRanges = wb.DefinedNames.InvalidNamedRanges().ToList();
        var localValidRanges = ws1.DefinedNames.ValidNamedRanges().ToList();
        var localInvalidRanges = ws1.DefinedNames.InvalidNamedRanges().ToList();

        var xlDefinedNames = globalValidRanges.ToList();
        await Assert.That(xlDefinedNames.Count).IsEqualTo(1);
        await Assert.That(xlDefinedNames.First().Name).IsEqualTo("Named range 2");

        await Assert.That(globalInvalidRanges.Count).IsEqualTo(2);
        await Assert.That(globalInvalidRanges.First().Name).IsEqualTo("Named range 4");
        await Assert.That(globalInvalidRanges.Last().Name).IsEqualTo("Named range 5");

        await Assert.That(localValidRanges.Count).IsEqualTo(1);
        await Assert.That(localValidRanges.First().Name).IsEqualTo("Named range 1");

        await Assert.That(localInvalidRanges.Count).IsEqualTo(0);
    }

    [Test]
    public async Task CanRenameNamedRange()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet1");
        var dn1 = wb.DefinedNames.Add("TEST", "=0.1");

        await Assert.That(wb.DefinedNames.TryGetValue("TEST", out _)).IsTrue();
        await Assert.That(wb.DefinedNames.TryGetValue("TEST1", out _)).IsFalse();

        dn1.Name = "TEST1";

        await Assert.That(wb.DefinedNames.TryGetValue("TEST", out _)).IsFalse();
        await Assert.That(wb.DefinedNames.TryGetValue("TEST1", out _)).IsTrue();

        var dn2 = wb.DefinedNames.Add("TEST2", "=TEST1*2");

        ws1.Cell(1, 1).FormulaA1 = "TEST1";
        ws1.Cell(2, 1).FormulaA1 = "TEST1*10";
        ws1.Cell(3, 1).FormulaA1 = "TEST2";
        ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

        await Assert.That((double)ws1.Cell(1, 1).Value).IsEqualTo(0.1).Within(XLHelper.Epsilon);
        await Assert.That((double)ws1.Cell(2, 1).Value).IsEqualTo(1.0).Within(XLHelper.Epsilon);
        await Assert.That((double)ws1.Cell(3, 1).Value).IsEqualTo(0.2).Within(XLHelper.Epsilon);
        await Assert.That((double)ws1.Cell(4, 1).Value).IsEqualTo(0.6).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task Can_save_and_load_defined_names()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet1 = wb.Worksheets.Add("Sheet1");
            var sheet2 = wb.Worksheets.Add("Sheet2");

            wb.DefinedNames.Add("wbNamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet1.DefinedNames.Add("sheet1NamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet2.DefinedNames.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var sheet1 = wb.Worksheet("Sheet1");
            var sheet2 = wb.Worksheet("Sheet2");

            await Assert.That(wb.DefinedNames.Count()).IsEqualTo(1);
            await Assert.That(wb.DefinedNames.Single().Name).IsEqualTo("wbNamedRange");
            await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            await Assert.That(wb.DefinedNames.Single().Ranges.Count).IsEqualTo(5);

            await Assert.That(sheet1.DefinedNames.Count()).IsEqualTo(1);
            await Assert.That(sheet1.DefinedNames.Single().Name).IsEqualTo("sheet1NamedRange");
            await Assert.That(sheet1.DefinedNames.Single().RefersTo).IsEqualTo("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            await Assert.That(sheet1.DefinedNames.Single().Ranges.Count).IsEqualTo(5);

            await Assert.That(sheet2.DefinedNames.Count()).IsEqualTo(1);
            await Assert.That(sheet2.DefinedNames.Single().Name).IsEqualTo("sheet2NamedRange");
            await Assert.That(sheet2.DefinedNames.Single().RefersTo).IsEqualTo("Sheet1!A1,Sheet2!A1");
            await Assert.That(sheet2.DefinedNames.Single().Ranges.Count).IsEqualTo(2);
        }
    }

    [Test]
    public async Task CopyNamedRangeDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var ws2 = wb.Worksheets.Add("Sheet2");
        var ranges = new XLRanges
        {
            ws1.Range("B2:E6"),
            ws2.Range("D1:E2")
        };
        var original = ws1.DefinedNames.Add("Named range", ranges);

        var copy = original.CopyTo(ws2);

        await Assert.That(ws1.DefinedNames.Count()).IsEqualTo(1);
        await Assert.That(ws2.DefinedNames.Count()).IsEqualTo(1);
        await Assert.That(original.Ranges.Count).IsEqualTo(2);
        await Assert.That(copy.Ranges.Count).IsEqualTo(2);
        await Assert.That(copy.Name).IsEqualTo(original.Name);
        await Assert.That(copy.Scope).IsEqualTo(original.Scope);
        await Assert.That(original.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet1!B2:E6");
        await Assert.That(original.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!D1:E2");
        await Assert.That(copy.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!D1:E2");
        await Assert.That(copy.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!B2:E6");
    }

    [Test]
    public async Task Copy_table_references_to_different_worksheet()
    {
        // When sheet-scoped name references a table and there is a table with same area in the
        // copied sheet, the copied defined name changes table reference to a new table. If
        // range differs, table reference is not modified.
        using var wb = new XLWorkbook();
        var orgSheet = wb.AddWorksheet();
        orgSheet.Cell("A1").InsertTable(["Data", "A", "B"], "OrgTable", true);
        orgSheet.Cell("C1").InsertTable(["Data", "A", "B"], "MiscTable", true);
        var originalName = orgSheet.DefinedNames.Add("TableName", "SUM(OrgTable[Data], MiscTable[Data])");

        var copySheet = wb.AddWorksheet();
        copySheet.Cell("A1").InsertTable(["Data", "A", "B"], "CopyTable", true);

        originalName.CopyTo(copySheet);

        var copyName = copySheet.DefinedNames.Single();
        await Assert.That(copyName.Name).IsEqualTo("TableName");
        await Assert.That(copyName.RefersTo).IsEqualTo("SUM(CopyTable[Data], MiscTable[Data])");
    }

    [Test]
    public async Task Copy_workbook_scoped_defined()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet");
        var name = wb.DefinedNames.Add("Name", "Sheet!$A$1");

        var copySheet = wb.AddWorksheet();
        var ex = await Assert.That(() => name.CopyTo(copySheet)).Throws<InvalidOperationException>()!;
        await Assert.That(ex.Message).IsEqualTo("Cannot copy workbook scoped defined name.");
    }

    [Test]
    public async Task Copy_defined_name_to_same_sheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Range("B2:E6").AddToNamed("Named range", XLScope.Worksheet);
        var dn = ws1.DefinedName("Named range");

        Action action = () => dn.CopyTo(ws1);

        await Assert.That(action).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task DeleteColumnUsedInNamedRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Column1");
        ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
        ws.FirstCell().CellRight(2).SetValue("Column3");
        ws.DefinedNames.Add("MyRange", "A1:C1");

        ws.Column(1).Delete();

        await Assert.That(ws.Cell("A1").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("B1").Value).IsEqualTo("Column3");
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task Formula_is_updated_on_sheet_rename()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Old name");
        var bookScopedName = wb.DefinedNames.Add("TEST", "ABS('Old name'!$B$5)");
        var sheetScopedName = ws.DefinedNames.Add("TEST1", "'Old name'!$D$7:$F$14");

        ws.Name = "Renamed";

        await Assert.That(bookScopedName.RefersTo).IsEqualTo("ABS(Renamed!$B$5)");
        await Assert.That(bookScopedName.Ranges.ToString()).IsEqualTo("Renamed!$B$5:$B$5");

        await Assert.That(sheetScopedName.RefersTo).IsEqualTo("Renamed!$D$7:$F$14");
        await Assert.That(sheetScopedName.Ranges.ToString()).IsEqualTo("Renamed!$D$7:$F$14");
    }

    [Test]
    public async Task MovingRanges()
    {
        var wb = new XLWorkbook();

        var sheet1 = wb.Worksheets.Add("Sheet1");
        var sheet2 = wb.Worksheets.Add("Sheet2");

        wb.DefinedNames.Add("wbNamedRange",
            "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
        sheet1.DefinedNames.Add("sheet1NamedRange",
            "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
        sheet2.DefinedNames.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

        sheet1.Row(1).InsertRowsAbove(2);
        sheet1.Row(1).Delete();
        sheet1.Column(1).InsertColumnsBefore(2);
        sheet1.Column(1).Delete();

        await Assert.That(wb.DefinedNames.First().RefersTo).IsEqualTo("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H");
        await Assert.That(sheet1.DefinedNames.First().RefersTo).IsEqualTo("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H");
        await Assert.That(sheet2.DefinedNames.First().RefersTo).IsEqualTo("Sheet1!B2,Sheet2!A1");

        // Were ForEach(dn => Assert...) under NUnit; ForEach takes an Action, so an awaited
        // assertion needs an explicit loop rather than an async lambda.
        foreach (var dn in wb.DefinedNames)
        {
            await Assert.That(dn.Scope).IsEqualTo(XLNamedRangeScope.Workbook);
        }

        foreach (var dn in sheet1.DefinedNames)
        {
            await Assert.That(dn.Scope).IsEqualTo(XLNamedRangeScope.Worksheet);
        }

        foreach (var dn in sheet2.DefinedNames)
        {
            await Assert.That(dn.Scope).IsEqualTo(XLNamedRangeScope.Worksheet);
        }
    }

    [Test]
    public async Task NamedRangeBecomesInvalidOnRangeAndWorksheetDeleting()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet 1");
        var ws2 = wb.Worksheets.Add("Sheet 2");
        ws1.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
        wb.DefinedNames.Add("Compound", new XLRanges
        {
            ws1.Range("C1:D2"),
            ws2.Range("A10:D15")
        });

        ws1.Rows(1, 5).Delete();
        ws1.Delete();

        await Assert.That(wb.DefinedNames.Count()).IsEqualTo(2);
        await Assert.That(wb.DefinedNames.ValidNamedRanges().Count()).IsEqualTo(0);

        // The row deletion reduces both Sheet 1 references to 'Sheet 1'!#REF!, and deleting the sheet
        // then drops the prefix, matching how a reference to a deleted sheet is stored elsewhere
        // (see NamedRangesFromDeletedSheetAreSavedWithoutAddress). Sheet 2 is untouched by either.
        await Assert.That(wb.DefinedNames.ElementAt(0).RefersTo).IsEqualTo("#REF!");
        await Assert.That(wb.DefinedNames.ElementAt(1).RefersTo).IsEqualTo("#REF!,'Sheet 2'!$A$10:$D$15");
    }

    [Test]
    public async Task NamedRangeBecomesInvalidOnRangeDeleting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet 1");
        ws.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
        wb.DefinedNames.Add("Compound", new XLRanges
        {
            ws.Range("C1:D2"),
            ws.Range("A10:D15")
        });

        ws.Rows(1, 5).Delete();

        await Assert.That(wb.DefinedNames.Count()).IsEqualTo(2);
        await Assert.That(wb.DefinedNames.ValidNamedRanges().Count()).IsEqualTo(0);
        // Simple is deleted outright; Compound loses C1:D2 and keeps A10:D15, shifted up five rows.
        await Assert.That(wb.DefinedNames.ElementAt(0).RefersTo).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(wb.DefinedNames.ElementAt(1).RefersTo).IsEqualTo("'Sheet 1'!#REF!,'Sheet 1'!$A$5:$D$10");
    }

    [Test]
    public async Task NamedRangeMayReferToExpression()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws1 = wb.AddWorksheet("Sheet1");
            wb.DefinedNames.Add("TEST", "=0.1");
            wb.DefinedNames.Add("TEST2", "=TEST*2");

            ws1.Cell(1, 1).FormulaA1 = "TEST";
            ws1.Cell(2, 1).FormulaA1 = "TEST*10";
            ws1.Cell(3, 1).FormulaA1 = "TEST2";
            ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

            await Assert.That((double)ws1.Cell(1, 1).Value).IsEqualTo(0.1).Within(XLHelper.Epsilon);
            await Assert.That((double)ws1.Cell(2, 1).Value).IsEqualTo(1.0).Within(XLHelper.Epsilon);
            await Assert.That((double)ws1.Cell(3, 1).Value).IsEqualTo(0.2).Within(XLHelper.Epsilon);
            await Assert.That((double)ws1.Cell(4, 1).Value).IsEqualTo(0.6).Within(XLHelper.Epsilon);

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws1 = wb.Worksheets.First();

            await Assert.That((double)ws1.Cell(1, 1).Value).IsEqualTo(0.1).Within(XLHelper.Epsilon);
            await Assert.That((double)ws1.Cell(2, 1).Value).IsEqualTo(1.0).Within(XLHelper.Epsilon);
            await Assert.That((double)ws1.Cell(3, 1).Value).IsEqualTo(0.2).Within(XLHelper.Epsilon);
            await Assert.That((double)ws1.Cell(4, 1).Value).IsEqualTo(0.6).Within(XLHelper.Epsilon);
        }
    }

    [Test]
    public async Task NamedRangeReferringToMultipleRangesCanBeSavedAndLoaded()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet 1");

            wb.DefinedNames.Add("Multirange named range", new XLRanges
            {
                ws.Range("A5:D5"),
                ws.Range("A15:D15")
            });

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.DefinedNames.Count()).IsEqualTo(1);
            var nr = (XLDefinedName)wb.DefinedNames.Single();
            await Assert.That(nr.RefersTo).IsEqualTo("'Sheet 1'!$A$5:$D$5,'Sheet 1'!$A$15:$D$15");
            await Assert.That(nr.Ranges.Count).IsEqualTo(2);
            await Assert.That(nr.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A5:D5");
            await Assert.That(nr.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A15:D15");
            var sheetRefs = nr.GetSheetReferencesList();
            await Assert.That(sheetRefs).Count().IsEqualTo(2);
            await Assert.That(sheetRefs[0]).IsEqualTo("'Sheet 1'!$A$5:$D$5");
            await Assert.That(sheetRefs[^1]).IsEqualTo("'Sheet 1'!$A$15:$D$15");
        }
    }

    [Test]
    public async Task Defined_names_referencing_sheet_range_become_invalid_when_sheet_is_deleted()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet 1");
        var ws2 = wb.Worksheets.Add("Sheet 2");
        var ws3 = wb.Worksheets.Add("Sheet'3");

        ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
        ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
        ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
        ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
        wb.DefinedNames.Add("Named range 5", new XLRanges
        {
            ws1.Range("A5:D5"),
            ws3.Range("A5:D5")
        });

        ws2.Delete();
        ws3.Delete();

        await Assert.That(ws1.DefinedNames.Count()).IsEqualTo(1);
        await Assert.That(ws1.DefinedNames.First().Name).IsEqualTo("Named range 1");
        await Assert.That(ws1.DefinedNames.First().Scope).IsEqualTo(XLNamedRangeScope.Worksheet);
        await Assert.That(ws1.DefinedNames.First().RefersTo).IsEqualTo("'Sheet 1'!$A$1:$D$1");
        await Assert.That(ws1.DefinedNames.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A1:D1");

        await Assert.That(wb.DefinedNames.Count()).IsEqualTo(3);

        await Assert.That(wb.DefinedNames.ElementAt(0).Name).IsEqualTo("Named range 2");
        await Assert.That(wb.DefinedNames.ElementAt(0).Scope).IsEqualTo(XLNamedRangeScope.Workbook);
        await Assert.That(wb.DefinedNames.ElementAt(0).RefersTo).IsEqualTo("'Sheet 1'!$A$2:$D$2");
        await Assert.That(wb.DefinedNames.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A2:D2");

        await Assert.That(wb.DefinedNames.ElementAt(1).Name).IsEqualTo("Named range 4");
        await Assert.That(wb.DefinedNames.ElementAt(1).Scope).IsEqualTo(XLNamedRangeScope.Workbook);
        await Assert.That(wb.DefinedNames.ElementAt(1).RefersTo).IsEqualTo("#REF!");
        await Assert.That(wb.DefinedNames.ElementAt(1).Ranges.Count).IsEqualTo(0);

        await Assert.That(wb.DefinedNames.ElementAt(2).Name).IsEqualTo("Named range 5");
        await Assert.That(wb.DefinedNames.ElementAt(2).Scope).IsEqualTo(XLNamedRangeScope.Workbook);
        await Assert.That(wb.DefinedNames.ElementAt(2).RefersTo).IsEqualTo("'Sheet 1'!$A$5:$D$5,#REF!");
        await Assert.That(wb.DefinedNames.ElementAt(2).Ranges.Count).IsEqualTo(1);
        await Assert.That(wb.DefinedNames.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A5:D5");
    }

    [Test]
    public async Task NamedRangesFromDeletedSheetAreSavedWithoutAddress()
    {
        // Range address referring to the deleted sheet look like #REF!A1:B2.
        // But workbooks with such references in named ranges Excel considers as broken files.
        // It requires #REF!

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet 1");
            var ws2 = wb.Worksheets.Add("Sheet 2");
            ws2.Range("A4:D4").AddToNamed("Test named range", XLScope.Workbook);
            ws2.Delete();
            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
        }
    }

    [Test]
    public async Task Only_worksheet_scoped_defined_names_are_copied_when_sheet_is_copied()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet1");
        ws1.FirstCell().InsertData(Enumerable.Range(1, 10));
        wb.DefinedNames.Add("wbNamedRange", ws1.Range("A1:A10"));
        ws1.DefinedNames.Add("wsNamedRange", ws1.Range("A3"));

        var ws2 = wb.AddWorksheet("Sheet2");
        ws2.FirstCell().InsertData(Enumerable.Range(101, 10));
        ws1.DefinedNames.Add("wsNamedRangeAcrossSheets", ws2.Range("A4"));

        ws1.Cell("C1").FormulaA1 = "=wbNamedRange";
        ws1.Cell("C2").FormulaA1 = "=wsNamedRange";
        ws1.Cell("C3").FormulaA1 = "=wsNamedRangeAcrossSheets";

        await Assert.That(ws1.Cell("C1").Value).IsEqualTo(1);
        await Assert.That(ws1.Cell("C2").Value).IsEqualTo(3);
        await Assert.That(ws1.Cell("C3").Value).IsEqualTo(104);

        var wsCopy = ws1.CopyTo("Copy");
        await Assert.That(wsCopy.Cell("C1").Value).IsEqualTo(1);
        await Assert.That(wsCopy.Cell("C2").Value).IsEqualTo(3);
        await Assert.That(wsCopy.Cell("C3").Value).IsEqualTo(104);

        await Assert.That(wb.DefinedName("wbNamedRange")!.Ranges.First().RangeAddress.ToStringRelative(true)).IsEqualTo("Sheet1!A1:A10");
        await Assert.That(wsCopy.DefinedName("wsNamedRange").Ranges.First().RangeAddress.ToStringRelative(true)).IsEqualTo("Copy!A3:A3");
        await Assert.That(wsCopy.DefinedName("wsNamedRangeAcrossSheets").Ranges.First().RangeAddress.ToStringRelative(true)).IsEqualTo("Sheet2!A4:A4");
    }

    [Test]
    public async Task Saved_defined_names_become_invalid_on_sheet_deleting()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws1 = wb.Worksheets.Add("Sheet 1");
            var ws2 = wb.Worksheets.Add("Sheet2");
            var ws3 = wb.Worksheets.Add("Sheet'3");

            ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
            ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
            ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
            ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
            wb.DefinedNames.Add("Named range 5", new XLRanges
            {
                ws1.Range("A5:D5"),
                ws3.Range("A5:D5")
            });

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet("Sheet2").Delete();
            wb.Worksheet("Sheet'3").Delete();
            wb.Save();
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws1 = wb.Worksheet("Sheet 1");
            await Assert.That(ws1.DefinedNames.Count()).IsEqualTo(1);
            await Assert.That(ws1.DefinedNames.First().Name).IsEqualTo("Named range 1");
            await Assert.That(ws1.DefinedNames.First().Scope).IsEqualTo(XLNamedRangeScope.Worksheet);
            await Assert.That(ws1.DefinedNames.First().RefersTo).IsEqualTo("'Sheet 1'!$A$1:$D$1");
            await Assert.That(ws1.DefinedNames.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A1:D1");

            await Assert.That(wb.DefinedNames.Count()).IsEqualTo(3);

            await Assert.That(wb.DefinedNames.ElementAt(0).Name).IsEqualTo("Named range 2");
            await Assert.That(wb.DefinedNames.ElementAt(0).Scope).IsEqualTo(XLNamedRangeScope.Workbook);
            await Assert.That(wb.DefinedNames.ElementAt(0).RefersTo).IsEqualTo("'Sheet 1'!$A$2:$D$2");
            await Assert.That(wb.DefinedNames.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A2:D2");

            await Assert.That(wb.DefinedNames.ElementAt(1).Name).IsEqualTo("Named range 4");
            await Assert.That(wb.DefinedNames.ElementAt(1).Scope).IsEqualTo(XLNamedRangeScope.Workbook);
            await Assert.That(wb.DefinedNames.ElementAt(1).RefersTo).IsEqualTo("#REF!");
            await Assert.That(wb.DefinedNames.ElementAt(1).Ranges.Count).IsEqualTo(0);

            await Assert.That(wb.DefinedNames.ElementAt(2).Name).IsEqualTo("Named range 5");
            await Assert.That(wb.DefinedNames.ElementAt(2).Scope).IsEqualTo(XLNamedRangeScope.Workbook);
            await Assert.That(wb.DefinedNames.ElementAt(2).RefersTo).IsEqualTo("'Sheet 1'!$A$5:$D$5,#REF!");
            await Assert.That(wb.DefinedNames.ElementAt(2).Ranges.Count).IsEqualTo(1);
            await Assert.That(wb.DefinedNames.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!A5:D5");
        }
    }

    [Test]
    public async Task TestInvalidNamedRangeOnWorkbookScope()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Column1");
        ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
        ws.FirstCell().CellRight(2).SetValue("Column3");

        await Assert.That(() => wb.DefinedNames.Add("MyRange", "A1:C1")).Throws<ArgumentException>();
    }

    [Test]
    public async Task WbContainsWsNamedRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

        await Assert.That(wb.DefinedNames.Contains("Sheet1!Name")).IsTrue();
        await Assert.That(wb.DefinedNames.Contains("Sheet1!NameX")).IsFalse();

        await Assert.That(wb.DefinedName("Sheet1!Name")).IsNotNull();
        await Assert.That(wb.DefinedName("Sheet1!NameX")).IsNull();

        var found1 = wb.DefinedNames.TryGetValue("Sheet1!Name", out var definedName1);
        await Assert.That(found1).IsTrue();
        await Assert.That(definedName1).IsNotNull();
        await Assert.That(definedName1!.Scope).IsEqualTo(XLNamedRangeScope.Worksheet);

        var found2 = wb.DefinedNames.TryGetValue("Sheet1!NameX", out var definedName2);
        await Assert.That(found2).IsFalse();
        await Assert.That(definedName2).IsNull();
    }

    [Test]
    public async Task WorkbookContainsNamedRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddToNamed("Name");

        await Assert.That(wb.DefinedNames.Contains("Name")).IsTrue();
        await Assert.That(wb.DefinedNames.Contains("NameX")).IsFalse();

        await Assert.That(wb.DefinedName("Name")).IsNotNull();
        await Assert.That(wb.DefinedName("NameX")).IsNull();

        var found1 = wb.DefinedNames.TryGetValue("Name", out var definedName1);
        await Assert.That(found1).IsTrue();
        await Assert.That(definedName1).IsNotNull();

        var found2 = wb.DefinedNames.TryGetValue("NameX", out var definedName2);
        await Assert.That(found2).IsFalse();
        await Assert.That(definedName2).IsNull();
    }

    [Test]
    public async Task WorksheetContainsNamedRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

        await Assert.That(ws.DefinedNames.Contains("Name")).IsTrue();
        await Assert.That(ws.DefinedNames.Contains("NameX")).IsFalse();

        await Assert.That(ws.DefinedName("Name")).IsNotNull();
        await Assert.That(() => ws.DefinedName("NameX")).Throws<KeyNotFoundException>();

        var found1 = ws.DefinedNames.TryGetValue("Name", out var definedName1);
        await Assert.That(found1).IsTrue();
        await Assert.That(definedName1).IsNotNull();

        var found2 = ws.DefinedNames.TryGetValue("NameX", out var definedName2);
        await Assert.That(found2).IsFalse();
        await Assert.That(definedName2).IsNull();
    }

    [Test]
    public async Task NamedRangeWithSameNameAsAFunction()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var a1 = ws.FirstCell();
        var a2 = a1.CellBelow();

        a1.SetValue(5).AddToNamed("RAND");
        a2.FormulaA1 = "=RAND * 10";

        await Assert.That(a2.GetDouble()).IsEqualTo(50);
    }

    [Test]
    public async Task DefinedName_SheetNameLikeCellRef_PreservesQuotes_OnRoundTrip()
    {
        // Excel requires it to be quoted in formulas.
        using var ms = new MemoryStream();

        // Create a workbook with a sheet "C05A" and a defined name referencing it.
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("C05A");
            ws.Cell("A1").Value = 1;
            // Directly set formula with quoted sheet name (as Excel would)
            wb.DefinedNames.Add("TestName", "'C05A'!$A$1:$A$10");
            wb.SaveAs(ms);
        }

        // Round-trip: open and save
        ms.Position = 0;
        using var ms2 = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            var dn = wb.DefinedNames.DefinedName("TestName");
            // The formula should still have quotes around C05A
            await Assert.That(dn.RefersTo).IsEqualTo("'C05A'!$A$1:$A$10");
            wb.SaveAs(ms2);
        }

        // Verify the saved XML still has quotes
        ms2.Position = 0;
        using (var wb = new XLWorkbook(ms2))
        {
            var dn = wb.DefinedNames.DefinedName("TestName");
            await Assert.That(dn.RefersTo).IsEqualTo("'C05A'!$A$1:$A$10");
        }
    }

    [Test]
    public async Task DefinedName_SheetNameLikeCellRef_AddFromRange_EscapesCorrectly()
    {
        // When creating a defined name from a range (not a formula string),
        // a sheet name that could be ambiguous with cell references should be quoted.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("C05A");
            ws.Cell("A1").Value = 1;
            ws.Range("A1:A10").AddToNamed("TestName", XLScope.Workbook);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var dn = wb.DefinedNames.DefinedName("TestName");
            // Sheet name C05A should be quoted because it looks like a cell reference prefix
            await Assert.That(dn.RefersTo).IsEqualTo("'C05A'!$A$1:$A$10");
        }
    }
}
