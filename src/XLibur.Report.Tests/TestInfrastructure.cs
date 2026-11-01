using System.Globalization;

// The report suite shares process-wide state with the workbooks it generates — temp files, the
// font engine and the current culture — so it runs serially, matching XLibur.Tests.
[assembly: NotInParallel]

namespace XLibur.Report.Tests;

/// <summary>Assembly-wide test defaults.</summary>
public static class TestDefaults
{
    /// <summary>The culture the suite runs under.</summary>
    internal static readonly CultureInfo DefaultCulture = CultureInfo.GetCultureInfo("en-US");

    [Before(HookType.Assembly)]
    public static void GlobalSetup() => SetCulture(DefaultCulture);

    [BeforeEvery(HookType.Test)]
    public static void ApplyCulture(TestContext context) => SetCulture(DefaultCulture);

    /// <summary>
    /// Only <see cref="CultureInfo.DefaultThreadCurrentCulture"/> is assigned, never
    /// <see cref="CultureInfo.CurrentCulture"/> — assigning the latter pins the value in an
    /// AsyncLocal that does not flow into the context TUnit runs the test body in. Safe because
    /// the assembly runs serially.
    /// </summary>
    internal static void SetCulture(CultureInfo culture)
    {
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
    }
}
