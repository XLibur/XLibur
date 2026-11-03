using System.Text;
using XLibur.Excel;

namespace XLibur.Operations.Tests;

public class ConvertAndReportTests
{
    private static readonly ConvertOperation Convert = new();

    [Test]
    public async Task CSV_to_xlsx_to_CSV_keeps_quoted_fields_numbers_and_text()
    {
        using var root = new TempDirectory();
        root.File("in.csv", "Name,Qty,Note\r\n\"Smith, J\",3,\"said \"\"hi\"\"\"\r\nLee,0.5,007\r\n");
        var context = Fixtures.Context(root);

        var toXlsx = Fixtures.Run(context, Convert, new ConvertRequest { Input = "in.csv", Output = "out.xlsx", Sheet = "Data" }).Outcome;
        var back = Fixtures.Run(context, Convert, new ConvertRequest { Input = "out.xlsx", Output = "back.csv" }).Outcome;

        await Assert.That(toXlsx.Error).IsNull();
        await Assert.That(back.Error).IsNull();
        using (var wb = new XLWorkbook(root.Combine("out.xlsx")))
        {
            await Assert.That(wb.Worksheet("Data").Cell("B3").Value).IsEqualTo((XLCellValue)0.5);
            await Assert.That(wb.Worksheet("Data").Cell("C3").Value).IsEqualTo((XLCellValue)7d);
        }

        await Assert.That(await File.ReadAllTextAsync(root.Combine("back.csv")))
            .IsEqualTo("Name,Qty,Note\r\n\"Smith, J\",3,\"said \"\"hi\"\"\"\r\nLee,0.5,7\r\n");
    }

    [Test]
    public async Task A_workbook_exports_to_JSON_rows_or_header_keyed_objects_with_calculated_values()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        var context = Fixtures.Context(root);

        Fixtures.Run(context, Convert, new ConvertRequest { Input = "sales.xlsx", Output = "rows.json" });
        Fixtures.Run(context, Convert, new ConvertRequest { Input = "sales.xlsx", Output = "objects.json", Header = true });

        var rows = await File.ReadAllTextAsync(root.Combine("rows.json"));
        var objects = await File.ReadAllTextAsync(root.Combine("objects.json"));
        await Assert.That(rows).StartsWith("[[\"Product\",\"Qty\",\"Price\",\"Total\"],[\"Apple\",3,1.5,4.5]");
        await Assert.That(objects).Contains("{\"Product\":\"Apple\",\"Qty\":3,\"Price\":1.5,\"Total\":4.5}");
    }

    [Test]
    public async Task JSON_objects_become_a_header_row_and_TSV_escapes_survive()
    {
        using var root = new TempDirectory();
        root.File("in.json", "[{\"a\":1,\"b\":\"x\\ty\"},{\"a\":true,\"c\":{\"error\":\"#N/A\"}}]");
        var context = Fixtures.Context(root);

        Fixtures.Run(context, Convert, new ConvertRequest { Input = "in.json", Output = "out.tsv" });

        await Assert.That(await File.ReadAllTextAsync(root.Combine("out.tsv"))).IsEqualTo("a\tb\tc\r\n1\tx\\ty\t\r\nTRUE\t\t#N/A\r\n");
    }

    [Test]
    [Arguments("out.csv")]
    [Arguments("out.tsv")]
    public async Task Export_prefixes_formula_like_text_with_a_quote_but_not_numbers_or_plain_text(string output)
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "in.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "=1+1";
            ws.Cell("A2").Value = -5;
            ws.Cell("A3").Value = "plain";
            ws.Cell("A4").FormulaA1 = "\"=\"&\"1+1\"";
            ws.Cell("A5").Value = "@SUM(A1)";
            ws.Cell("A6").Value = "-x";
            ws.Cell("A7").Value = "+x";
        });
        var context = Fixtures.Context(root);

        var outcome = Fixtures.Run(context, Convert, new ConvertRequest { Input = "in.xlsx", Output = output, Recalculate = true }).Outcome;

        await Assert.That(outcome.Error).IsNull();
        await Assert.That(await File.ReadAllTextAsync(root.Combine(output)))
            .IsEqualTo("'=1+1\r\n-5\r\nplain\r\n'=1+1\r\n'@SUM(A1)\r\n'-x\r\n'+x\r\n");
    }

    [Test]
    public async Task Conversion_refuses_to_replace_an_output_without_overwrite_and_xlsx_to_xlsm()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        root.File("taken.csv", "old");
        var context = Fixtures.Context(root);

        var exists = Fixtures.Run(context, Convert, new ConvertRequest { Input = "sales.xlsx", Output = "taken.csv" }).Outcome;
        var macro = Fixtures.Run(context, Convert, new ConvertRequest { Input = "sales.xlsx", Output = "sales.xlsm" }).Outcome;

        await Assert.That(exists.Error!.Code).IsEqualTo("file_exists");
        await Assert.That(macro.Error!.Kind).IsEqualTo(ErrorKind.Unsupported);
    }

    [Test]
    public async Task A_report_template_binds_JSON_data_and_lists_template_errors()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "template.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Report");
            ws.Cell("A1").Value = "{{ Company }}";
            ws.Cell("B1").Value = "{{ Due }}";
            ws.Cell("C1").Value = "{{ Code }}";
            ws.Cell("A3").Value = "{{ item.Product }}";
            ws.Cell("B3").Value = "{{ item.Qty }}";
            ws.Cell("C3").Value = "{{ item.Qty * }}";
            wb.DefinedNames.Add("Items", "Report!$A$3:$C$3");
        });
        root.File("data.json", "{\"Company\":\"Contoso\",\"Due\":{\"date\":\"2026-01-31\"},\"Code\":\"2026-02-01\",\"Items\":[{\"Product\":\"Apple\",\"Qty\":3},{\"Product\":\"Pear\",\"Qty\":5}]}");

        var (outcome, text) = Fixtures.Run(Fixtures.Context(root), new RenderReportOperation(),
            new RenderReportRequest { Template = "template.xlsx", Data = "data.json", Output = "report.xlsx" });

        var result = (RenderReportResult)outcome.Result!;
        using var report = new XLWorkbook(root.Combine("report.xlsx"));
        var ws = report.Worksheet("Report");
        await Assert.That(ws.Cell("A1").Value).IsEqualTo((XLCellValue)"Contoso");
        await Assert.That(ws.Cell("B1").Value).IsEqualTo((XLCellValue)new DateTime(2026, 1, 31));
        await Assert.That(ws.Cell("C1").Value).IsEqualTo((XLCellValue)"2026-02-01");
        await Assert.That(ws.Cell("A3").Value).IsEqualTo((XLCellValue)"Apple");
        await Assert.That(ws.Cell("A4").Value).IsEqualTo((XLCellValue)"Pear");
        await Assert.That(ws.Cell("B4").Value).IsEqualTo((XLCellValue)5d);
        await Assert.That(result.Errors.Count).IsGreaterThan(0);
        await Assert.That(outcome.ExitCode).IsEqualTo(1);
        await Assert.That(text).Contains("Report!C3");
    }

    [Test]
    public async Task Report_data_must_be_one_JSON_object()
    {
        using var root = new TempDirectory();
        Fixtures.Sales(root);
        root.File("data.json", "[1,2]");

        var outcome = Fixtures.Run(Fixtures.Context(root), new RenderReportOperation(),
            new RenderReportRequest { Template = "sales.xlsx", Data = "data.json", Output = "report.xlsx" }).Outcome;

        await Assert.That(outcome.Error!.Code).IsEqualTo("bad_data");
        await Assert.That(File.Exists(root.Combine("report.xlsx"))).IsFalse();
    }

    [Test]
    public async Task Export_refuses_a_grid_over_the_cell_limit_and_writes_nothing()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "far.xlsx", wb => wb.AddWorksheet("S").Cell("J11").Value = 1);
        var context = Fixtures.Context(root, limits: Limits.Default with { MaxGridCells = 100 });

        var outcome = Fixtures.Run(context, Convert, new ConvertRequest { Input = "far.xlsx", Output = "far.csv" }).Outcome;

        await Assert.That(outcome.Error!.Code).IsEqualTo("grid_too_large");
        await Assert.That(File.Exists(root.Combine("far.csv"))).IsFalse();
    }

    [Test]
    public async Task Export_of_one_cell_at_XFD1048576_is_refused_by_the_default_limit()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "corner.xlsx", wb => wb.AddWorksheet("S").Cell("XFD1048576").Value = 1);

        var outcome = Fixtures.Run(Fixtures.Context(root), Convert, new ConvertRequest { Input = "corner.xlsx", Output = "corner.json" }).Outcome;

        await Assert.That(outcome.Error!.Code).IsEqualTo("grid_too_large");
    }

    [Test]
    public async Task Header_keyed_JSON_refuses_duplicate_headers_and_keys_columns_past_the_header()
    {
        using var root = new TempDirectory();
        root.File("dup.csv", "Name,Name\r\nAlice,Bob\r\n");
        root.File("wide.csv", "Name\r\nAlice,30\r\n");
        var context = Fixtures.Context(root);

        var duplicate = Fixtures.Run(context, Convert, new ConvertRequest { Input = "dup.csv", Output = "dup.json", Header = true }).Outcome;
        var wide = Fixtures.Run(context, Convert, new ConvertRequest { Input = "wide.csv", Output = "wide.json", Header = true }).Outcome;

        await Assert.That(duplicate.Error!.Code).IsEqualTo("duplicate_header");
        await Assert.That(File.Exists(root.Combine("dup.json"))).IsFalse();
        await Assert.That(wide.Error).IsNull();
        await Assert.That(await File.ReadAllTextAsync(root.Combine("wide.json"))).IsEqualTo("[{\"Name\":\"Alice\",\"column2\":30}]");
    }

    [Test]
    public async Task CSV_export_quotes_text_that_would_import_as_another_type()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "in.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = "007";
            ws.Cell("B1").Value = "TRUE";
            ws.Cell("C1").Value = "1e3";
            ws.Cell("D1").Value = 7;
            ws.Cell("E1").Value = true;
        });
        var context = Fixtures.Context(root);

        Fixtures.Run(context, Convert, new ConvertRequest { Input = "in.xlsx", Output = "out.csv" });
        Fixtures.Run(context, Convert, new ConvertRequest { Input = "out.csv", Output = "back.xlsx" });

        await Assert.That(await File.ReadAllTextAsync(root.Combine("out.csv"))).IsEqualTo("\"007\",\"TRUE\",\"1e3\",7,TRUE\r\n");
        using var back = new XLWorkbook(root.Combine("back.xlsx"));
        var row = back.Worksheet("Sheet1").Row(1);
        await Assert.That(row.Cell(1).Value).IsEqualTo((XLCellValue)"007");
        await Assert.That(row.Cell(2).Value).IsEqualTo((XLCellValue)"TRUE");
        await Assert.That(row.Cell(4).Value).IsEqualTo((XLCellValue)7d);
        await Assert.That(row.Cell(5).Value.IsBoolean && row.Cell(5).Value.GetBoolean()).IsTrue();
    }

    [Test]
    public async Task A_last_CSV_record_of_one_quoted_empty_field_is_kept_without_a_trailing_newline()
    {
        using var root = new TempDirectory();
        root.File("in.csv", "a\r\n\"\"");

        Fixtures.Run(Fixtures.Context(root), Convert, new ConvertRequest { Input = "in.csv", Output = "out.json" });

        await Assert.That(await File.ReadAllTextAsync(root.Combine("out.json"))).IsEqualTo("[[\"a\"],[\"\"]]");
    }

    [Test]
    public async Task Text_output_is_UTF8_without_a_byte_order_mark()
    {
        using var root = new TempDirectory();
        Fixtures.Build(root, "u.xlsx", wb => wb.AddWorksheet("S").Cell("A1").Value = "Zürich €");

        Fixtures.Run(Fixtures.Context(root), Convert, new ConvertRequest { Input = "u.xlsx", Output = "u.csv" });

        var bytes = await File.ReadAllBytesAsync(root.Combine("u.csv"));
        await Assert.That(bytes[0]).IsNotEqualTo((byte)0xEF);
        await Assert.That(Encoding.UTF8.GetString(bytes)).IsEqualTo("Zürich €\r\n");
    }
}
