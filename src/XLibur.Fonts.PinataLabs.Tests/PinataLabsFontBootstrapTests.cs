using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Fonts.Tests;
using XLibur.Graphics;

namespace XLibur.Fonts.PinataLabs.Tests;

/// <summary>
/// Tests for registering the PinataLabs engine as the default font engine.
/// These mutate the process-wide <see cref="LoadOptions.DefaultFontEngine"/>, so the fixture is
/// non-parallelizable and each test restores the previous value.
/// </summary>
[NotInParallel]
public class PinataLabsFontBootstrapTests
{
    private IXLFontEngine? _previousDefault;

    [Before(HookType.Test)]
    public void SaveDefault() => _previousDefault = LoadOptions.DefaultFontEngine;

    [After(HookType.Test)]
    public void RestoreDefault() => LoadOptions.DefaultFontEngine = _previousDefault;

    [Test]
    public async Task CreateDefault_ReturnsWorkingEngine()
    {
        var engine = PinataLabsFontBootstrap.CreateDefault();

        var width = engine.GetTextWidth("Hello", new DummyFont("Arial", 11), 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task CreateDefault_UnknownFont_FallsBackAndMeasures()
    {
        var engine = PinataLabsFontBootstrap.CreateDefault();

        // A font that exists on no machine forces resolution down to the named fallback, or the embedded
        // CarlitoBare where that is missing too (the path exercised on headless Linux with no system fonts).
        var width = engine.GetTextWidth("Hello", new DummyFont("TotallyFakeNonExistentFont12345", 11), 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task CreateDefault_ReturnsTheSharedInstance()
    {
        await Assert.That(PinataLabsFontBootstrap.CreateDefault()).IsSameReferenceAs(PinataLabsFontEngine.Instance.Value);
    }

    [Test]
    public async Task Register_SetsDefaultFontEngineWhenNoneRegistered()
    {
        LoadOptions.DefaultFontEngine = null;

        PinataLabsFontBootstrap.Register();

        await Assert.That(LoadOptions.DefaultFontEngine).IsTypeOf<PinataLabsFontEngine>();
    }

    [Test]
    public async Task Register_DoesNotOverrideAnExistingDefault()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var existing = PinataLabsFontEngine.CreateOnlyWithFonts(fallbackStream);
        LoadOptions.DefaultFontEngine = existing;

        PinataLabsFontBootstrap.Register();

        await Assert.That(LoadOptions.DefaultFontEngine).IsSameReferenceAs(existing);
    }

    [Test]
    public async Task Registered_Workbook_UsesEngineForAdjustToContents()
    {
        LoadOptions.DefaultFontEngine = null;
        PinataLabsFontBootstrap.Register();

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Hello World";
        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

    private sealed class DummyFont : IXLFontBase
    {
        public DummyFont(string name, double size)
        {
            FontName = name;
            FontSize = size;
        }

        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strikethrough { get; set; }
        public XLFontUnderlineValues Underline { get; set; } = XLFontUnderlineValues.None;
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public bool Shadow { get; set; }
        public XLColor FontColor { get; set; } = XLColor.Black;
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; } = XLFontFamilyNumberingValues.NotApplicable;
        public XLFontCharSet FontCharSet { get; set; } = XLFontCharSet.Default;
        public XLFontScheme FontScheme { get; set; }
    }
}
