using System.Collections.Generic;
using System.IO;
using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Tests that structured table references with colons in column names are preserved
/// during formula rewriting and save/load round trips.
/// </summary>
[TestFixture]
internal class StructuredReferenceColonTests
{
    [Test]
    public void Formula_with_colon_in_column_name_is_preserved_on_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var ws2 = wb.AddWorksheet("Sheet2");
        var data = new List<object[]>
        {
            new object[] { "Week Date", "Some Header: Other" },
            new object[] { "2026-01-01", "1" }
        };
        var range = ws2.FirstCell().InsertData(data);
        range.CreateTable("Table2");

        ws.FirstCell().Value = "2026-01-01";
        ws.Cell("B1").FormulaA1 =
            "=_xlfn.XLOOKUP(A1,Table2[Week Date],Table2[Some Header: Other])";

        var formula = ws.Cell("B1").FormulaA1;
        Assert.That(formula, Does.Contain("Table2[Some Header: Other]"),
            "Structured reference with colon in column name must be preserved");
        Assert.That(formula, Does.Not.Contain("#REF!"),
            "Formula must not contain #REF! error");
    }

    [Test]
    public void Formula_with_colon_in_column_name_survives_save_and_load()
    {
        using var ms = new MemoryStream();

        // Save
        using (var wb = new XLWorkbook())
        {
            var ws2 = wb.AddWorksheet("Sheet2");
            var data = new List<object[]>
            {
                new object[] { "Week Date", "Some Header: Other" },
                new object[] { "2026-01-01", "1" }
            };
            var range = ws2.FirstCell().InsertData(data);
            range.CreateTable("Table2");

            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().Value = "2026-01-01";
            ws.Cell("B1").FormulaA1 =
                "=_xlfn.XLOOKUP(A1,Table2[Week Date],Table2[Some Header: Other])";

            wb.SaveAs(ms);
        }

        // Load and verify
        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var formula = wb.Worksheet("Sheet1").Cell("B1").FormulaA1;
            Assert.That(formula, Does.Contain("Table2[Some Header: Other]"),
                "Structured reference with colon in column name must survive save/load round trip");
            Assert.That(formula, Does.Not.Contain("#REF!"));
        }
    }

    [Test]
    public void Normal_range_formula_is_not_affected()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("C1").FormulaA1 = "=SUM(A1:B1)";

        Assert.That(ws.Cell("C1").FormulaA1, Is.EqualTo("SUM(A1:B1)"));
        Assert.That(ws.Cell("C1").Value, Is.EqualTo(3));
    }

    [Test]
    public void Structured_reference_without_colon_still_works()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var ws2 = wb.AddWorksheet("Sheet2");
        var data = new List<object[]>
        {
            new object[] { "Week Date", "Value" },
            new object[] { "2026-01-01", "42" }
        };
        var range = ws2.FirstCell().InsertData(data);
        range.CreateTable("Table2");

        ws.Cell("A1").FormulaA1 = "=Table2[Value]";

        var formula = ws.Cell("A1").FormulaA1;
        Assert.That(formula, Does.Contain("Table2[Value]"));
        Assert.That(formula, Does.Not.Contain("#REF!"));
    }

    [Test]
    public void Multiple_structured_references_in_one_formula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var ws2 = wb.AddWorksheet("Sheet2");
        var data = new List<object[]>
        {
            new object[] { "Key", "Col: A", "Col: B" },
            new object[] { "x", "1", "2" }
        };
        var range = ws2.FirstCell().InsertData(data);
        range.CreateTable("MyTable");

        ws.Cell("A1").FormulaA1 =
            "=_xlfn.XLOOKUP(\"x\",MyTable[Key],MyTable[Col: A])+_xlfn.XLOOKUP(\"x\",MyTable[Key],MyTable[Col: B])";

        var formula = ws.Cell("A1").FormulaA1;
        Assert.That(formula, Does.Contain("MyTable[Col: A]"),
            "First structured reference with colon must be preserved");
        Assert.That(formula, Does.Contain("MyTable[Col: B]"),
            "Second structured reference with colon must be preserved");
        Assert.That(formula, Does.Not.Contain("#REF!"));
    }

    [Test]
    public void Multiple_headers_with_colons()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var ws2 = wb.AddWorksheet("Sheet2");
        var data = new List<object[]>
        {
            new object[] { "ID", "Start: Date", "End: Date" },
            new object[] { "1", "2026-01-01", "2026-12-31" }
        };
        var range = ws2.FirstCell().InsertData(data);
        range.CreateTable("Dates");

        ws.Cell("A1").FormulaA1 =
            "=_xlfn.XLOOKUP(1,Dates[ID],Dates[Start: Date])";
        ws.Cell("B1").FormulaA1 =
            "=_xlfn.XLOOKUP(1,Dates[ID],Dates[End: Date])";

        Assert.That(ws.Cell("A1").FormulaA1, Does.Contain("Dates[Start: Date]"));
        Assert.That(ws.Cell("B1").FormulaA1, Does.Contain("Dates[End: Date]"));
        Assert.That(ws.Cell("A1").FormulaA1, Does.Not.Contain("#REF!"));
        Assert.That(ws.Cell("B1").FormulaA1, Does.Not.Contain("#REF!"));
    }
}
