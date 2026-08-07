using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests;

/// <summary>
/// Guards the boundary spec 13 drew: XLibur.Report depends on XLibur as a published package, so
/// what it can reach has to be a surface XLibur promises to keep.
/// </summary>
public class PublicSurfaceTests
{
    /// <summary>
    /// The types spec 13 chose <em>not</em> to publish, having replaced Report's use of them with
    /// <c>XLFunctionLibrary</c> and the <c>IXLPivotCache</c> source members.
    /// </summary>
    /// <remarks>
    /// These are deep calc-engine and coordinate representations. Publishing any of them freezes an
    /// implementation detail permanently, and would have been done for the sake of two call sites.
    /// If a change makes one of these public, the design is being bypassed — revise the spec rather
    /// than the test.
    /// </remarks>
    private static readonly string[] MustStayInternal =
    [
        "XLibur.Excel.CalcEngine.XLCalcEngine",
        "XLibur.Excel.CalcEngine.FunctionRegistry",
        "XLibur.Excel.CalcEngine.FunctionDefinition",
        "XLibur.Excel.CalcEngine.CalcContext",
        "XLibur.Excel.CalcEngine.AnyValue",
        "XLibur.Excel.CalcEngine.ScalarValue",
        "XLibur.Excel.CalcEngine.Exceptions.MissingContextException",
        "XLibur.Excel.XLPivotCache",
        "XLibur.Excel.XLPivotSourceReference",
        "XLibur.Excel.IXLPivotSource",
        "XLibur.Excel.Coordinates.SheetArea",
        "XLibur.Excel.Coordinates.Area",

        // The style key/value pair for alignment. Every other component pair - border, fill, font,
        // number format, protection, colour, and the composite style itself - has always been
        // internal; these two were public with nothing depending on it, no public member handing one
        // out, and no entry in the list below. They were made internal to match, and belong here so
        // the asymmetry cannot come back unnoticed.
        "XLibur.Excel.XLAlignmentKey",
        "XLibur.Excel.XLAlignmentValue",
    ];

    [Test]
    public async Task The_types_report_no_longer_reaches_for_are_still_internal()
    {
        var assembly = typeof(XLWorkbook).Assembly;
        var leaked = new List<string>();

        foreach (var name in MustStayInternal)
        {
            var type = assembly.GetType(name, throwOnError: false);

            // A rename is not a failure of this test — the point is that nothing here is public.
            if (type is not null && type.IsVisible)
                leaked.Add(name);
        }

        await Assert.That(leaked).IsEmpty();
    }

    [Test]
    public async Task The_surface_report_depends_on_is_public()
    {
        var assembly = typeof(XLWorkbook).Assembly;

        var functionLibrary = assembly.GetType("XLibur.Excel.CalcEngine.XLFunctionLibrary", throwOnError: false);
        await Assert.That(functionLibrary).IsNotNull();
        await Assert.That(functionLibrary!.IsVisible).IsTrue();

        var noContext = assembly.GetType(
            "XLibur.Excel.CalcEngine.Exceptions.XLNoWorksheetContextException",
            throwOnError: false);
        await Assert.That(noContext).IsNotNull();
        await Assert.That(noContext!.IsVisible).IsTrue();

        var sourceKind = assembly.GetType("XLibur.Excel.XLPivotSourceKind", throwOnError: false);
        await Assert.That(sourceKind).IsNotNull();
        await Assert.That(sourceKind!.IsVisible).IsTrue();

        var expected = new[] { "SourceKind", "SourceRange", "SourceName", "SourceWorksheet", "SetSourceRange" };
        var members = typeof(IXLPivotCache).GetMembers().Select(m => m.Name).ToList();

        foreach (var member in expected)
            await Assert.That(members).Contains(member);
    }

    /// <summary>
    /// The friend grant spec 13 removed. Its test-assembly counterpart is deliberately kept — see
    /// the comment in <c>XLibur/Properties/AssemblyInfo.cs</c>.
    /// </summary>
    [Test]
    public async Task XLibur_Report_is_not_a_friend_assembly()
    {
        var grants = typeof(XLWorkbook).Assembly
            .GetCustomAttributes(typeof(System.Runtime.CompilerServices.InternalsVisibleToAttribute), false)
            .Cast<System.Runtime.CompilerServices.InternalsVisibleToAttribute>()
            .Select(a => a.AssemblyName)
            .ToList();

        await Assert.That(grants).DoesNotContain("XLibur.Report");
        await Assert.That(grants).Contains("XLibur.Report.Tests");
    }
}
