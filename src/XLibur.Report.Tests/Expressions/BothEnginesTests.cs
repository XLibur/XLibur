using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.DynamicLinq;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tests.Expressions;

/// <summary>
/// The claim the compatibility package makes: a template's <em>structure</em> is engine-independent, so
/// pointing an upstream-authored template at <see cref="DynamicLinqExpressionEngine"/> is all that is
/// needed to run it.
/// </summary>
/// <remarks>
/// <para>
/// Each test here builds the same report twice — the same defined name, the same options row, the same
/// tags, the same formatting — differing only in how the expressions are written, and asserts the two
/// generated workbooks agree cell for cell. That is a stronger statement than "the compat engine works":
/// it says the engine seam is where the whole of the difference lives, which is the property that lets
/// the package plug in and out.
/// </para>
/// <para>
/// The Scriban side doubles as a control. If a structural feature broke for both engines these tests
/// would still pass, which is why the feature suites elsewhere assert values rather than agreement.
/// </para>
/// </remarks>
public class BothEnginesTests
{
    private static List<SaleItem> Items() => new()
    {
        new SaleItem { Product = "Trowel", Region = "North", Category = "Retail", Quantity = 96, UnitPrice = 4.20m },
        new SaleItem { Product = "Hoe", Region = "South", Category = "Trade", Quantity = 4, UnitPrice = 240.00m },
        new SaleItem { Product = "Twine", Region = "North", Category = "Retail", Quantity = 240, UnitPrice = 1.15m },
        new SaleItem { Product = "Cloche", Region = "South", Category = "Trade", Quantity = 18, UnitPrice = 42.00m },
    };

    /// <summary>
    /// Builds a template, generates it under <paramref name="engine"/>, and hands back the sheet.
    /// </summary>
    private static (XLWorkbook Workbook, XLGenerateResult Result) Generate(
        Action<IXLWorksheet> build,
        IExpressionEngine? engine,
        Action<IXLTemplate>? data = null)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        build(sheet);

        using var template = new XLTemplate(workbook, engine);
        template.AddVariable("Company", "Contoso");
        template.AddVariable("Items", Items());
        data?.Invoke(template);

        return (workbook, template.Generate());
    }

    /// <summary>
    /// Generates <paramref name="scriban"/> under the default engine and <paramref name="upstream"/>
    /// under the compatibility engine, and asserts the two agree over the used range.
    /// </summary>
    private static async Task Agree(Action<IXLWorksheet> scriban, Action<IXLWorksheet> upstream)
    {
        var (defaultWorkbook, defaultResult) = Generate(scriban, engine: null);
        using var _ = defaultWorkbook;

        var (compatWorkbook, compatResult) = Generate(upstream, new DynamicLinqExpressionEngine());
        using var __ = compatWorkbook;

        await Assert.That(defaultResult.HasErrors).IsFalse();
        await Assert.That(compatResult.HasErrors).IsFalse();

        var expected = defaultWorkbook.Worksheet("Report");
        var actual = compatWorkbook.Worksheet("Report");

        var used = expected.RangeUsed(XLCellsUsedOptions.All)!.RangeAddress;

        await Assert.That(actual.RangeUsed(XLCellsUsedOptions.All)!.RangeAddress.ToString())
            .IsEqualTo(used.ToString());

        for (var row = used.FirstAddress.RowNumber; row <= used.LastAddress.RowNumber; row++)
        {
            for (var column = used.FirstAddress.ColumnNumber; column <= used.LastAddress.ColumnNumber; column++)
            {
                await Assert.That(actual.Cell(row, column).GetFormattedString())
                    .IsEqualTo(expected.Cell(row, column).GetFormattedString());
            }
        }
    }

    /// <summary>Headings in row 1, the repeated row in 2, the options row in 3.</summary>
    private static void Frame(IXLWorksheet sheet, string title, params string[] repeated)
    {
        sheet.Cell("A1").Value = title;

        for (var i = 0; i < repeated.Length; i++)
        {
            sheet.Cell(2, i + 1).Value = repeated[i];
        }

        sheet.Workbook.DefinedNames.Add("Items", sheet.Range(2, 1, 3, Math.Max(1, repeated.Length)));
    }

    [Test]
    public async Task ABoundRangeAgrees()
    {
        await Agree(
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Product }}", "{{ item.Quantity }}"),
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Product }}", "{{ item.Quantity }}"));
    }

    /// <summary>
    /// The same report written the two ways round: a Scriban filter against a .NET method call.
    /// </summary>
    [Test]
    public async Task TheTwoSyntaxesForTheSameThingAgree()
    {
        await Agree(
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Product | string.upcase }}", "{{ item.Quantity }}"),
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Product.ToUpper() }}", "{{ item.Quantity }}"));
    }

    /// <summary>A computed value, arrived at by each engine's own arithmetic.</summary>
    [Test]
    public async Task ArithmeticAgrees()
    {
        await Agree(
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Product }}", "{{ item.Quantity * item.UnitPrice }}"),
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Product }}", "{{ item.Quantity * item.UnitPrice }}"));
    }

    /// <summary>
    /// Mixed text and expressions, which goes through <c>Interpolate</c> rather than <c>Evaluate</c> —
    /// a separate path in both engines.
    /// </summary>
    [Test]
    public async Task InterpolatedTextAgrees()
    {
        await Agree(
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Quantity }} x {{ item.Product }}"),
            sheet => Frame(sheet, "{{ Company }}", "{{ item.Quantity }} x {{ item.Product }}"));
    }

    /// <summary>
    /// Tags are engine-independent by construction — they are read from cell text before any expression
    /// is evaluated — so a sorted, grouped, totalled report has to come out the same under either.
    /// </summary>
    [Test]
    public async Task GroupingSortingAndTotallingAgree()
    {
        void Build(IXLWorksheet sheet)
        {
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "Product";
            sheet.Cell("C1").Value = "Total";

            sheet.Cell("A2").Value = "{{ item.Region }}";
            sheet.Cell("B2").Value = "{{ item.Product }}";
            sheet.Cell("C2").Value = "{{ item.Total }}";

            sheet.Cell("A3").Value = "<<Group merge>>";
            sheet.Cell("B3").Value = "<<Sort>>";
            sheet.Cell("C3").Value = "<<Sum>>";

            sheet.Workbook.DefinedNames.Add("Items", sheet.Range("A2:C3"));
        }

        await Agree(Build, Build);
    }

    /// <summary>
    /// <c>&lt;&lt;If&gt;&gt;</c> evaluates its <c>test</c> through the engine, so the two syntaxes for
    /// the same question have to filter the same rows.
    /// </summary>
    [Test]
    public async Task AConditionalRowAgrees()
    {
        await Agree(
            sheet =>
            {
                Frame(sheet, "{{ Company }}", "{{ item.Product }}", "{{ item.Quantity }}");
                sheet.Cell("C3").Value = "<<If test=\"item.Quantity > 10\">>";
            },
            sheet =>
            {
                Frame(sheet, "{{ Company }}", "{{ item.Product }}", "{{ item.Quantity }}");
                sheet.Cell("C3").Value = "<<If test=\"item.Quantity > 10\">>";
            });
    }

    /// <summary>
    /// The <c>&amp;=</c> prefix builds a formula from interpolated text, so it exercises the engine at a
    /// point where a locale-dependent number would produce a broken formula rather than a wrong value.
    /// </summary>
    [Test]
    public async Task AGeneratedFormulaAgrees()
    {
        void Build(IXLWorksheet sheet)
        {
            sheet.Cell("A1").Value = "Product";
            sheet.Cell("A2").Value = "{{ item.Product }}";
            sheet.Cell("B2").Value = "&=SUM(1, {{ item.Quantity }})";
            sheet.Workbook.DefinedNames.Add("Items", sheet.Range("A2:B3"));
        }

        await Agree(Build, Build);
    }

    /// <summary>
    /// Upstream templates reach a workbook variable from inside a range with an <c>@</c> prefix. The
    /// default engine has no such form, so this one is asserted against a value rather than by agreement.
    /// </summary>
    [Test]
    public async Task AnAtPrefixedGlobalWorksInsideARange()
    {
        var (workbook, result) = Generate(
            sheet => Frame(sheet, "Report", "{{ @Company }}", "{{ item.Product }}"),
            new DynamicLinqExpressionEngine());

        using var _ = workbook;

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(workbook.Worksheet("Report").Cell("A2").GetFormattedString()).IsEqualTo("Contoso");
        await Assert.That(workbook.Worksheet("Report").Cell("A5").GetFormattedString()).IsEqualTo("Contoso");
    }

    /// <summary>
    /// LINQ over the bound collection, which is the upstream idiom the default engine replaces with the
    /// Excel-function bridge. Asserted against a value, there being no equivalent to agree with.
    /// </summary>
    [Test]
    public async Task LinqOverTheCollectionWorksInATemplate()
    {
        var (workbook, result) = Generate(
            sheet =>
            {
                sheet.Cell("A1").Value = "{{ Items.Sum(x => x.Total) }}";
                sheet.Cell("A3").Value = "{{ item.Product }}";
                sheet.Workbook.DefinedNames.Add("Items", sheet.Range("A3:A4"));
            },
            new DynamicLinqExpressionEngine());

        using var _ = workbook;

        var expected = Items().Sum(item => item.Total);

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetNumber()).IsEqualTo((double)expected);
    }

    /// <summary>
    /// A typed result reaches the cell as a number under the compatibility engine too — the property
    /// that lets a total be formatted and summed rather than being text that looks like one.
    /// </summary>
    [Test]
    public async Task ASingleExpressionKeepsItsType()
    {
        var (workbook, result) = Generate(
            sheet => Frame(sheet, "Report", "{{ item.Product }}", "{{ item.Total }}"),
            new DynamicLinqExpressionEngine());

        using var _ = workbook;

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(workbook.Worksheet("Report").Cell("B2").Value.IsNumber).IsTrue();
    }

    /// <summary>
    /// A bad expression is reported rather than thrown under this engine as well, and the rest of the
    /// report is still generated — the guarantee is the engine seam's, not one engine's.
    /// </summary>
    [Test]
    public async Task ABadExpressionIsReportedNotThrown()
    {
        var (workbook, result) = Generate(
            sheet =>
            {
                sheet.Cell("A1").Value = "{{ item.Quantity + }}";
                sheet.Cell("A3").Value = "{{ item.Product }}";
                sheet.Workbook.DefinedNames.Add("Items", sheet.Range("A3:A4"));
            },
            new DynamicLinqExpressionEngine());

        using var _ = workbook;

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(workbook.Worksheet("Report").Cell("A3").GetFormattedString()).IsEqualTo("Trowel");
    }

    /// <summary>
    /// The Excel-function bridge is a default-engine feature. Under the compatibility engine the call is
    /// reported as an unknown name rather than silently producing nothing, and generation continues.
    /// </summary>
    [Test]
    public async Task TheExcelFunctionBridgeIsAbsentUnderTheCompatibilityEngine()
    {
        var (workbook, result) = Generate(
            sheet =>
            {
                sheet.Cell("A1").Value = "{{ SUM(1, 2) }}";
                sheet.Cell("A3").Value = "{{ item.Product }}";
                sheet.Workbook.DefinedNames.Add("Items", sheet.Range("A3:A4"));
            },
            new DynamicLinqExpressionEngine());

        using var _ = workbook;

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(workbook.Worksheet("Report").Cell("A3").GetFormattedString()).IsEqualTo("Trowel");
    }

    /// <summary>
    /// A whole report under the compatibility engine still passes the OpenXML validator on save: the
    /// engine changes what cells contain, and nothing else.
    /// </summary>
    [Test]
    public async Task AReportGeneratedUnderTheCompatibilityEnginePassesTheValidator()
    {
        var (workbook, result) = Generate(
            sheet =>
            {
                sheet.Cell("A1").Value = "{{ @Company }}";
                sheet.Cell("A3").Value = "{{ item.Region }}";
                sheet.Cell("B3").Value = "{{ item.Product.ToUpper() }}";
                sheet.Cell("C3").Value = "{{ item.Total }}";
                sheet.Cell("A4").Value = "<<Group>>";
                sheet.Cell("C4").Value = "<<Sum>>";
                sheet.Workbook.DefinedNames.Add("Items", sheet.Range("A3:C4"));
            },
            new DynamicLinqExpressionEngine());

        using var _ = workbook;

        await Assert.That(result.HasErrors).IsFalse();

        using var stream = new MemoryStream();
        await Assert.That(() => workbook.SaveAs(stream, validate: true)).ThrowsNothing();
    }
}
