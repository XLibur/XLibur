using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using XLibur.Excel;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Extensions;
/// <summary>
/// <see cref="XmlWriterExtensions"/> formats numbers into a reusable buffer instead of allocating
/// a string per value. Those writes end up in the saved file verbatim, so the output has to stay
/// byte-identical to <see cref="ObjectExtensions.ToInvariantString{T}"/>, which is what the writer
/// used before and what round-trip stability depends on.
/// </summary>
public class XmlWriterExtensionsTests
{
    [Test]
    [Arguments(0d)]
    [Arguments(1d)]
    [Arguments(-1d)]
    [Arguments(0.1d)]
    [Arguments(-0.5d)]
    [Arguments(1234567890.123456d)]
    [Arguments(1e-300)]
    [Arguments(1e300)]
    [Arguments(-1e-300)]
    [Arguments(-1e300)]
    [Arguments(double.Epsilon)]
    [Arguments(double.MaxValue)]
    [Arguments(double.MinValue)]
    [Arguments(1d / 3d)]
    [Arguments(2d / 3d)]
    [Arguments(0.1d + 0.2d)]
    public async Task WriteNumberValue_double_matches_ToInvariantString(double value)
    {
        await Assert.That(WriteDouble(value)).IsEqualTo(value.ToInvariantString());
    }

    [Test]
    public async Task WriteNumberValue_double_matches_ToInvariantString_for_random_values()
    {
#pragma warning disable S2245 // Deterministic seed keeps failures reproducible
        var random = new Random(1234);
#pragma warning restore S2245

        for (var i = 0; i < 20_000; i++)
        {
            var value = NextDouble(random);
            await Assert.That(WriteDouble(value)).IsEqualTo(value.ToInvariantString()).Because($"value #{i}");
        }
    }

    [Test]
    public async Task WriteNumberValue_serial_dates_match_ToInvariantString()
    {
#pragma warning disable S2245 // Deterministic seed keeps failures reproducible
        var random = new Random(4321);
#pragma warning restore S2245

        var baseDate = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        for (var i = 0; i < 20_000; i++)
        {
            var date = baseDate
                .AddDays(random.Next(0, 60_000))
                .AddSeconds(random.Next(0, 86_400))
                .AddMilliseconds(random.Next(0, 1000));

            var serial = date.ToSerialDateTime();
            await Assert.That(WriteDouble(serial)).IsEqualTo(serial.ToInvariantString()).Because($"date #{i}: {date:O}");
        }
    }

    [Test]
    [Arguments(0)]
    [Arguments(1)]
    [Arguments(-1)]
    [Arguments(12345)]
    [Arguments(int.MaxValue)]
    [Arguments(int.MinValue)]
    public async Task WriteNumberValue_int_matches_ToInvariantString(int value)
    {
        await Assert.That(WriteInt(value)).IsEqualTo(value.ToInvariantString());
    }

    [Test]
    [Arguments(0u)]
    [Arguments(1u)]
    [Arguments(12345u)]
    [Arguments(uint.MaxValue)]
    public async Task WriteNumberValue_uint_matches_ToInvariantString(uint value)
    {
        await Assert.That(WriteUInt(value)).IsEqualTo(value.ToInvariantString());
    }

    [Test]
    public async Task WriteNumberValue_int_matches_ToInvariantString_for_random_values()
    {
#pragma warning disable S2245 // Deterministic seed keeps failures reproducible
        var random = new Random(99);
#pragma warning restore S2245

        for (var i = 0; i < 20_000; i++)
        {
            var value = random.Next(int.MinValue, int.MaxValue);
            await Assert.That(WriteInt(value)).IsEqualTo(value.ToInvariantString()).Because($"value #{i}");
        }
    }

    /// <summary>
    /// The number buffer is thread-static and reused across calls, so a second value must not be
    /// contaminated by the leftovers of a longer first one.
    /// </summary>
    [Test]
    public async Task WriteNumberValue_reuses_the_buffer_without_leaking_previous_digits()
    {
        // "G15", not round-trip: 15 significant digits, matching ToInvariantString.
        await Assert.That(WriteDouble(double.MinValue)).IsEqualTo("-1.79769313486232E+308");
        await Assert.That(WriteDouble(1d)).IsEqualTo("1");
        await Assert.That(WriteInt(int.MinValue)).IsEqualTo("-2147483648");
        await Assert.That(WriteInt(7)).IsEqualTo("7");
    }

    private static string WriteDouble(double value) => Capture(w => w.WriteNumberValue(value));

    private static string WriteInt(int value) => Capture(w => w.WriteNumberValue(value));

    private static string WriteUInt(uint value) => Capture(w => w.WriteNumberValue(value));

    /// <summary>
    /// Capture what the extension writes as element content, which is exactly how the sheet writer
    /// emits cell values.
    /// </summary>
    private static string Capture(Action<XmlWriter> write)
    {
        var sb = new StringBuilder();
        var settings = new XmlWriterSettings
        {
            OmitXmlDeclaration = true,
            ConformanceLevel = ConformanceLevel.Fragment,
        };

        using (var writer = XmlWriter.Create(sb, settings))
        {
            writer.WriteStartElement("v");
            write(writer);
            writer.WriteEndElement();
        }

        var xml = sb.ToString();
        return xml["<v>".Length..^"</v>".Length];
    }

    private static double NextDouble(Random random)
    {
        return (random.Next(8)) switch
        {
            0 => random.NextDouble(),
            1 => random.NextDouble() * 10_000,
            2 => random.NextDouble() * 1e12,
            3 => random.NextDouble() * 1e-12,
            4 => -random.NextDouble() * 1e6,
            5 => random.Next(-1_000_000, 1_000_000),
            6 => Math.Round(random.NextDouble() * 10_000, 2),
            _ => BitConverter.Int64BitsToDouble(NextFiniteBits(random)),
        };
    }

    /// <summary>Random bit patterns, excluding NaN/Infinity which never reach the writer.</summary>
    private static long NextFiniteBits(Random random)
    {
        while (true)
        {
            var bits = ((long)random.Next() << 32) | (uint)random.Next();
            var candidate = BitConverter.Int64BitsToDouble(bits);
            if (!double.IsNaN(candidate) && !double.IsInfinity(candidate))
                return bits;
        }
    }
}
