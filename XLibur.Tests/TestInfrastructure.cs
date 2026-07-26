// The test project does not enable nullable reference types, but this file annotates the two
// members below. Opting just this file in keeps the annotations legal (CS8632) without turning
// null analysis on for the whole suite.
#nullable enable

using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using TUnit.Core.Interfaces;

// The suite was written against NUnit, which does not parallelise by default. TUnit does.
// These tests share a lot of process-wide state (temp files, the calc engine, the font
// engine registered in LoadOptions.DefaultFontEngine, and the current culture, which the
// [SetCulture] tests mutate), so the assembly is pinned to serial execution to preserve
// the original semantics. Revisit as a separate exercise once the migration has settled.
[assembly: NotInParallel]

namespace XLibur.Tests;

/// <summary>
/// Assembly-wide test defaults. Replaces the old <c>[SetUpFixture]</c>/<c>[OneTimeSetUp]</c>
/// pair, which TUnit models as an assembly-level hook instead.
/// </summary>
public static class TestDefaults
{
    [Before(HookType.Assembly)]
    public static void GlobalSetup()
    {
        Fonts.SixLabors.V1.SixLaborsV1FontBootstrap.Register();
        SetCulture(DefaultCulture);
    }

    /// <summary>The culture the suite runs under unless a test opts out via [SetCulture].</summary>
    internal static readonly CultureInfo DefaultCulture = CultureInfo.GetCultureInfo("en-US");

    /// <summary>
    /// Applies the culture for each test: a <see cref="SetCultureAttribute"/> on the test
    /// method, else one on its class, else the assembly default.
    /// </summary>
    /// <remarks>
    /// Resolved here rather than from the attribute itself because TUnit fires event
    /// receivers for both the method-level and class-level attribute, and the class-level
    /// one won -- so a method asking for cs-CZ inside a class declaring en-US ran as en-US.
    /// Doing the lookup in one place also resets culture between tests, undoing the several
    /// tests that assign <c>Thread.CurrentThread.CurrentCulture</c> directly and would
    /// otherwise leak into the next test on the same pooled thread.
    /// </remarks>
    [BeforeEvery(HookType.Test)]
    public static void ApplyCulture(TestContext context)
    {
        var culture = ResolveCulture(context) ?? DefaultCulture;

        SetCulture(culture);
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;
    }

    private static CultureInfo? ResolveCulture(TestContext context)
    {
        var details = context.Metadata.TestDetails;
        var classType = details.ClassType;

        var onMethod = classType
            .GetMethods()
            .Where(m => m.Name == details.MethodName)
            .Select(m => m.GetCustomAttribute<SetCultureAttribute>())
            .FirstOrDefault(a => a is not null);
        if (onMethod is not null)
        {
            return CultureInfo.GetCultureInfo(onMethod.Culture);
        }

        var onClass = classType.GetCustomAttribute<SetCultureAttribute>(inherit: true);
        return onClass is null ? null : CultureInfo.GetCultureInfo(onClass.Culture);
    }

    /// <summary>
    /// Only <see cref="CultureInfo.DefaultThreadCurrentCulture"/> is assigned, never
    /// <see cref="CultureInfo.CurrentCulture"/>. Assigning CurrentCulture pins the value in
    /// an AsyncLocal for that execution context, which both stops later
    /// DefaultThreadCurrentCulture changes from being observed and fails to flow into the
    /// context TUnit runs the test body in. Setting only the process-wide default keeps the
    /// switch visible everywhere -- safe because the assembly runs serially.
    /// </summary>
    internal static void SetCulture(CultureInfo culture)
    {
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
    }
}

/// <summary>
/// Bridges the equality gap between NUnit and TUnit for <see cref="XLibur.Excel.XLCellValue"/>.
/// </summary>
/// <remarks>
/// Many calc-engine tests declare their expected value as <c>object</c> and compare it
/// against an <c>XLCellValue</c>. NUnit's comparer discovered the <c>IEquatable&lt;double&gt;</c>,
/// <c>IEquatable&lt;XLError&gt;</c>, ... implementations on XLCellValue and dispatched to them.
/// TUnit compares the declared types, and <c>XLCellValue.Equals(object)</c> only matches
/// another XLCellValue, so a boxed <c>double</c> or <c>XLError</c> never compares equal.
/// Converting the expected value first restores the original semantics.
/// </remarks>
public static class ExpectedCellValue
{
    public static XLibur.Excel.XLCellValue From(object? expected) => expected switch
    {
        null => XLibur.Excel.Blank.Value,
        XLibur.Excel.XLCellValue cellValue => cellValue,
        XLibur.Excel.Blank blank => blank,
        XLibur.Excel.XLError error => error,
        bool logical => logical,
        string text => text,
        DateTime dateTime => dateTime,
        TimeSpan timeSpan => timeSpan,
        _ => Convert.ToDouble(expected, CultureInfo.InvariantCulture),
    };
}

/// <summary>
/// Stand-in for NUnit's <c>[SetCulture]</c>, which TUnit has no equivalent for.
/// A declarative marker; TestDefaults.ApplyCulture reads it before each test. Safe because the assembly runs tests serially (see
/// <c>[assembly: NotInParallel]</c> above); culture is thread state, so this would be
/// racy if tests ran concurrently.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false)]
public sealed class SetCultureAttribute : Attribute
{
    public SetCultureAttribute(string culture) => Culture = culture;

    public string Culture { get; }
}

/// <summary>
/// Stand-in for NUnit's <c>[Culture]</c>, which restricts a test to runs under a given
/// culture rather than setting one. The assembly pins en-US in
/// <see cref="TestDefaults.GlobalSetup"/> and every use of this attribute asks for en-US,
/// so the constraint is always satisfied and this is a marker only. It is kept so the
/// original intent stays visible at the call site.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = true)]
public sealed class CultureAttribute : Attribute
{
    public CultureAttribute(string culture) => Culture = culture;

    public string Culture { get; }
}
