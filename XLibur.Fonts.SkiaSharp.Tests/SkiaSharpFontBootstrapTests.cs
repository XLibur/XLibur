using System.Collections.Generic;
using System.IO;
using XLibur.Excel;
using XLibur.Graphics;
using System.Threading.Tasks;

namespace XLibur.Fonts.SkiaSharp.Tests;
/// <summary>
/// Tests for auto-registration of the SkiaSharp engine as the default font engine.
/// These mutate the process-wide <see cref="LoadOptions.DefaultFontEngine"/>, so the fixture is
/// non-parallelizable and each test restores the previous value.
/// </summary>
[NotInParallel]
public class SkiaSharpFontBootstrapTests
{
    private IXLFontEngine? _previousDefault;

    [Before(HookType.Test)]
    public void SaveDefault() => _previousDefault = LoadOptions.DefaultFontEngine;

    [After(HookType.Test)]
    public void RestoreDefault() => LoadOptions.DefaultFontEngine = _previousDefault;

    [Test]
    public async Task CreateDefault_ReturnsWorkingEngine()
    {
        var engine = SkiaSharpFontBootstrap.CreateDefault();

        var width = engine.GetTextWidth("Hello", new DummyFont("Arial", 11), 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task CreateDefault_UnknownFont_FallsBackAndMeasures()
    {
        var engine = SkiaSharpFontBootstrap.CreateDefault();

        // A font that exists on no machine forces resolution down to the embedded CarlitoBare fallback
        // (this is the path exercised on headless Linux/serverless with no system fonts).
        var width = engine.GetTextWidth("Hello", new DummyFont("TotallyFakeNonExistentFont12345", 11), 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task Register_SetsDefaultFontEngineWhenNoneRegistered()
    {
        LoadOptions.DefaultFontEngine = null;

        SkiaSharpFontBootstrap.Register();

        await Assert.That(LoadOptions.DefaultFontEngine).IsNotNull();
    }

    [Test]
    public async Task Register_DoesNotOverrideAnExistingDefault()
    {
        var existing = SkiaSharpFontBootstrap.CreateDefault();
        LoadOptions.DefaultFontEngine = existing;

        SkiaSharpFontBootstrap.Register();

        await Assert.That(LoadOptions.DefaultFontEngine).IsSameReferenceAs(existing);
    }

    [Test]
    public async Task ZeroConfig_Workbook_AutoRegistersDefaultEngineViaProbe()
    {
        // No explicit font engine anywhere: the core probe must reflectively discover this package.
        LoadOptions.DefaultFontEngine = null;

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Hello World";
        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

    [Test]
    public async Task ZeroConfig_InsertDataSample_JustWorks()
    {
        // Verbatim README-style sample: no font engine configured anywhere.
        LoadOptions.DefaultFontEngine = null;

        var data = new List<object[]>
        {
            new object[] { "Cheesecake", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 }
        };

        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Cool cheesecake stuff");
            ws.Cell("B3").InsertData(data);

            // Exercise the font engine (auto-fit measures text) and persist the file.
            ws.Columns().AdjustToContents();
            wb.SaveAs(ms);
        }

        // Round-trip: the saved file must reload with the inserted values intact.
        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var sheet = reloaded.Worksheet("Cool cheesecake stuff");

        using (Assert.Multiple())
        {
            await Assert.That(sheet.Cell("B3").GetString()).IsEqualTo("Cheesecake");
            await Assert.That(sheet.Cell("C3").GetValue<int>()).IsEqualTo(14);
            await Assert.That(sheet.Cell("B5").GetString()).IsEqualTo("Muffin");
            await Assert.That(sheet.Cell("C5").GetValue<int>()).IsEqualTo(10);

        }
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
