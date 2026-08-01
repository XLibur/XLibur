using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Tests.Excel.CalcEngine;

public class XLFunctionLibraryTests
{
    private static XLCellValue Invoke(XLFunctionLibrary library, string name, params XLCellValue[] arguments)
    {
        return library.TryInvoke(name, arguments.AsSpan(), out var result)
            ? result
            : throw new InvalidOperationException($"No function named '{name}'.");
    }

    [Test]
    public async Task Names_includes_the_common_functions()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(library.Names).IsNotEmpty();
        await Assert.That(library.Names).Contains("SUM");
        await Assert.That(library.Names).Contains("ROUND");
    }

    [Test]
    public async Task Names_are_all_invocable()
    {
        var library = new XLFunctionLibrary();
        var unresolved = new List<string>();

        // Every advertised name resolves — Names and TryInvoke read the same registry, and a caller
        // registering one adapter per name relies on that. Throwing counts as resolved: the call
        // was dispatched, it just needed a grid.
        foreach (var name in library.Names)
        {
            try
            {
                if (!library.TryInvoke(name, [], out _))
                    unresolved.Add(name);
            }
            catch (XLNoWorksheetContextException)
            {
            }
        }

        await Assert.That(unresolved).IsEmpty();
    }

    [Test]
    public async Task TryInvoke_calls_the_function()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(Invoke(library, "SUM", 1, 2, 3)).IsEqualTo((XLCellValue)6.0);
        await Assert.That(Invoke(library, "ROUND", 3.14159, 2)).IsEqualTo((XLCellValue)3.14);
        await Assert.That(Invoke(library, "UPPER", "abc")).IsEqualTo((XLCellValue)"ABC");
    }

    [Test]
    public async Task TryInvoke_matches_the_name_case_insensitively()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(Invoke(library, "sum", 1, 2)).IsEqualTo((XLCellValue)3.0);
        await Assert.That(Invoke(library, "Sum", 1, 2)).IsEqualTo((XLCellValue)3.0);
    }

    [Test]
    public async Task TryInvoke_returns_false_for_an_unknown_function()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(library.TryInvoke("NOT_A_FUNCTION", [1.0], out var result)).IsFalse();
        await Assert.That(result).IsEqualTo(default(XLCellValue));
    }

    [Test]
    public async Task TryInvoke_reports_wrong_arity_as_an_error_result()
    {
        var library = new XLFunctionLibrary();

        // The call was made and could not succeed, which Excel reports as a value — not as
        // "there is no such function", which is what false means.
        await Assert.That(library.TryInvoke("ROUND", [], out var tooFew)).IsTrue();
        await Assert.That(tooFew).IsEqualTo((XLCellValue)XLError.IncompatibleValue);

        await Assert.That(library.TryInvoke("PI", [1.0, 2.0, 3.0], out var tooMany)).IsTrue();
        await Assert.That(tooMany).IsEqualTo((XLCellValue)XLError.IncompatibleValue);
    }

    [Test]
    public async Task TryInvoke_reports_a_failed_call_as_an_error_result()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(Invoke(library, "CHAR", -2)).IsEqualTo((XLCellValue)XLError.IncompatibleValue);
    }

    [Test]
    public async Task TryInvoke_throws_for_a_function_that_needs_a_worksheet()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(() => library.TryInvoke("ROW", [], out _))
            .Throws<XLNoWorksheetContextException>();
    }

    [Test]
    public async Task TryInvoke_throws_for_a_null_name()
    {
        var library = new XLFunctionLibrary();

        await Assert.That(() => library.TryInvoke(null!, [], out _))
            .Throws<ArgumentNullException>();
    }

    [Test]
    public async Task TryInvoke_round_trips_every_scalar_kind()
    {
        var library = new XLFunctionLibrary();

        // IF returns its argument untouched, so it reports what survived the conversion in and out.
        await Assert.That(Invoke(library, "IF", true, 42.5)).IsEqualTo((XLCellValue)42.5);
        await Assert.That(Invoke(library, "IF", true, "text")).IsEqualTo((XLCellValue)"text");
        var logical = Invoke(library, "IF", true, true);
        await Assert.That(logical.IsBoolean).IsTrue();
        await Assert.That(logical.GetBoolean()).IsTrue();

        await Assert.That(Invoke(library, "IF", true, XLError.DivisionByZero))
            .IsEqualTo((XLCellValue)XLError.DivisionByZero);
    }

    [Test]
    public async Task TryInvoke_keeps_a_blank_result_blank()
    {
        var library = new XLFunctionLibrary();

        // Not 0. There is no cell here to apply the "a blank formula result is written as zero"
        // rule, so the caller is told what the function actually returned.
        var result = Invoke(library, "IF", true, Blank.Value);

        await Assert.That(result.IsBlank).IsTrue();
    }

    [Test]
    public async Task TryInvoke_reports_an_array_result_as_an_error()
    {
        var library = new XLFunctionLibrary();

        // SEQUENCE(2, 2) is a 2x2 array: no single value a cell could hold, and no XLCellValue that
        // could carry it. Reported the way Excel reports a value it cannot place.
        await Assert.That(Invoke(library, "SEQUENCE", 2, 2)).IsEqualTo((XLCellValue)XLError.IncompatibleValue);
    }

    [Test]
    public async Task Culture_defaults_to_invariant_and_is_honoured()
    {
        var czech = new XLFunctionLibrary(CultureInfo.GetCultureInfo("cs-CZ"));
        var invariant = new XLFunctionLibrary();

        // A decimal comma is a number in cs-CZ and not in the invariant culture.
        await Assert.That(Invoke(czech, "VALUE", "1,5")).IsEqualTo((XLCellValue)1.5);
        await Assert.That(Invoke(invariant, "VALUE", "1.5")).IsEqualTo((XLCellValue)1.5);
    }

    [Test]
    public async Task An_instance_is_safe_to_share_across_threads()
    {
        // XLibur.Report shares one library across every template precisely to avoid rebuilding the
        // function table, so concurrent TryInvoke has to hold up.
        var library = new XLFunctionLibrary();

        var results = await Task.WhenAll(Enumerable.Range(0, 64).Select(i => Task.Run(() =>
        {
            var sum = Invoke(library, "SUM", i, i);
            var text = Invoke(library, "UPPER", $"v{i}");
            return (Sum: sum, Text: text);
        })));

        for (var i = 0; i < results.Length; ++i)
        {
            await Assert.That(results[i].Sum).IsEqualTo((XLCellValue)(i * 2.0));
            await Assert.That(results[i].Text).IsEqualTo((XLCellValue)$"V{i}");
        }
    }
}
