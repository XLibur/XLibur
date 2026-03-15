using System;
using System.Drawing;
using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// Benchmarks for XLStyle key struct GetHashCode performance.
/// Tests whether replacing HashCode with manual hash computation improves throughput.
/// See: https://github.com/ClosedXML/ClosedXML/pull/2677
/// </summary>
[MemoryDiagnoser]
public class StyleKeyHashCodeBenchmarks
{
    private const int Iterations = 100_000;

    private XLFontKey[] _fontKeys = null!;
    private XLBorderKey[] _borderKeys = null!;
    private XLFillKey[] _fillKeys = null!;
    private XLColorKey[] _colorKeys = null!;
    private XLStyleKey[] _styleKeys = null!;

    [GlobalSetup]
    public void Setup()
    {
        var random = new Random(42);

        _colorKeys = new XLColorKey[Iterations];
        for (var i = 0; i < Iterations; i++)
        {
            _colorKeys[i] = (i % 3) switch
            {
                0 => new XLColorKey
                {
                    ColorType = XLColorType.Color,
                    Color = Color.FromArgb(random.Next(256), random.Next(256), random.Next(256))
                },
                1 => new XLColorKey
                {
                    ColorType = XLColorType.Theme,
                    ThemeColor = (XLThemeColor)(random.Next(12)),
                    ThemeTint = random.NextDouble()
                },
                _ => new XLColorKey
                {
                    ColorType = XLColorType.Indexed,
                    Indexed = random.Next(64)
                }
            };
        }

        _fontKeys = new XLFontKey[Iterations];
        for (var i = 0; i < Iterations; i++)
        {
            _fontKeys[i] = new XLFontKey
            {
                Bold = random.Next(2) == 1,
                Italic = random.Next(2) == 1,
                Underline = (XLFontUnderlineValues)(random.Next(5)),
                Strikethrough = random.Next(2) == 1,
                VerticalAlignment = (XLFontVerticalTextAlignmentValues)(random.Next(3)),
                Shadow = random.Next(2) == 1,
                FontSize = 8 + random.Next(20),
                FontColor = _colorKeys[i],
                FontName = random.Next(4) switch
                {
                    0 => "Calibri",
                    1 => "Arial",
                    2 => "Consolas",
                    _ => "Georgia"
                },
                FontFamilyNumbering = (XLFontFamilyNumberingValues)(random.Next(6)),
                FontCharSet = XLFontCharSet.Ansi,
                FontScheme = (XLFontScheme)(random.Next(3))
            };
        }

        _borderKeys = new XLBorderKey[Iterations];
        for (var i = 0; i < Iterations; i++)
        {
            _borderKeys[i] = new XLBorderKey
            {
                LeftBorder = (XLBorderStyleValues)(random.Next(14)),
                LeftBorderColor = _colorKeys[i],
                RightBorder = (XLBorderStyleValues)(random.Next(14)),
                RightBorderColor = _colorKeys[(i + 1) % Iterations],
                TopBorder = (XLBorderStyleValues)(random.Next(14)),
                TopBorderColor = _colorKeys[(i + 2) % Iterations],
                BottomBorder = (XLBorderStyleValues)(random.Next(14)),
                BottomBorderColor = _colorKeys[(i + 3) % Iterations],
                DiagonalBorder = (XLBorderStyleValues)(random.Next(14)),
                DiagonalBorderColor = _colorKeys[(i + 4) % Iterations],
                DiagonalUp = random.Next(2) == 1,
                DiagonalDown = random.Next(2) == 1
            };
        }

        _fillKeys = new XLFillKey[Iterations];
        for (var i = 0; i < Iterations; i++)
        {
            _fillKeys[i] = new XLFillKey
            {
                PatternType = (XLFillPatternValues)(random.Next(19)),
                BackgroundColor = _colorKeys[i],
                PatternColor = _colorKeys[(i + 1) % Iterations]
            };
        }

        _styleKeys = new XLStyleKey[Iterations];
        for (var i = 0; i < Iterations; i++)
        {
            _styleKeys[i] = new XLStyleKey
            {
                Alignment = new XLAlignmentKey
                {
                    Horizontal = (XLAlignmentHorizontalValues)(random.Next(8)),
                    Vertical = (XLAlignmentVerticalValues)(random.Next(5)),
                    Indent = random.Next(5),
                    JustifyLastLine = random.Next(2) == 1,
                    ReadingOrder = (XLAlignmentReadingOrderValues)(random.Next(3)),
                    RelativeIndent = random.Next(5),
                    ShrinkToFit = random.Next(2) == 1,
                    TextRotation = random.Next(180),
                    WrapText = random.Next(2) == 1
                },
                Border = _borderKeys[i],
                Fill = _fillKeys[i],
                Font = _fontKeys[i],
                IncludeQuotePrefix = random.Next(2) == 1,
                NumberFormat = new XLNumberFormatKey
                {
                    NumberFormatId = random.Next(50),
                    Format = random.Next(3) switch
                    {
                        0 => "",
                        1 => "#,##0.00",
                        _ => "yyyy-mm-dd"
                    }
                },
                Protection = new XLProtectionKey
                {
                    Locked = random.Next(2) == 1,
                    Hidden = random.Next(2) == 1
                }
            };
        }
    }

    [Benchmark]
    public int FontKey_GetHashCode()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum ^= _fontKeys[i].GetHashCode();
        return sum;
    }

    [Benchmark]
    public int BorderKey_GetHashCode()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum ^= _borderKeys[i].GetHashCode();
        return sum;
    }

    [Benchmark]
    public int FillKey_GetHashCode()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum ^= _fillKeys[i].GetHashCode();
        return sum;
    }

    [Benchmark]
    public int ColorKey_GetHashCode()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum ^= _colorKeys[i].GetHashCode();
        return sum;
    }

    [Benchmark]
    public int StyleKey_GetHashCode()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum ^= _styleKeys[i].GetHashCode();
        return sum;
    }
}
