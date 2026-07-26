using System.Globalization;

namespace XLibur.Fonts.SixLabors.Tests;

/// <summary>
/// Assembly-wide test defaults.
/// Replaces NUnit's <c>[assembly: SetCulture("en-US")]</c>, which TUnit has no direct
/// equivalent for. Culture is pinned on the thread pool rather than per test so that it
/// also applies to any threads TUnit creates for the run.
/// </summary>
public static class TestDefaults
{
    [Before(HookType.Assembly)]
    public static void PinCulture()
    {
        var enUs = CultureInfo.GetCultureInfo("en-US");
        CultureInfo.DefaultThreadCurrentCulture = enUs;
        CultureInfo.DefaultThreadCurrentUICulture = enUs;
        CultureInfo.CurrentCulture = enUs;
        CultureInfo.CurrentUICulture = enUs;
    }
}
