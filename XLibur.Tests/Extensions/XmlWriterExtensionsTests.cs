using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using NUnit.Framework;
using XLibur.Excel;
using XLibur.Extensions;

namespace XLibur.Tests.Extensions;

/// <summary>
/// <see cref="XmlWriterExtensions"/> formats numbers into a reusable buffer instead of allocating
/// a string per value. Those writes end up in the saved file verbatim, so the output has to stay
/// byte-identical to <see cref="ObjectExtensions.ToInvariantString{T}"/>, which is what the writer
/// used before and what round-trip stability depends on.
/// </summary>
[TestFixture]
public class XmlWriterExtensionsTests
{
    [TestCase(0d)]
    [TestCase(1d)]
    [TestCase(-1d)]
    [TestCase(0.1d)]
    [TestCase(-0.5d)]
    [TestCase(1234567890.123456d)]
    [TestCase(1e-300)]
    [TestCase(1e300)]
    [TestCase(-1e-300)]
    [TestCase(-1e300)]
    [TestCase(double.Epsilon)]
    [TestCase(double.MaxValue)]
    [TestCase(double.MinValue)]
    [TestCase(1d / 3d)]
    [TestCase(2d / 3d)]
    [TestCase(0.1d + 0.2d)]
    public void WriteNumberValue_double_matches_ToInvariantString(double value)
    {
        Assert.AreEqual(value.ToInvariantString(), WriteDouble(value));
    }

    [Test]
    public void WriteNumberValue_double_matches_ToInvariantString_for_random_values()
    {
#pragma warning disable S2245 // Deterministic seed keeps failures reproducible
        var random = new Random(1234);
#pragma warning restore S2245

        for (var i = 0; i < 20_000; i++)
        {
            var value = NextDouble(random);
            Assert.AreEqual(value.ToInvariantString(), WriteDouble(value), $"value #{i}");
        }
    }

    [Test]
    public void WriteNumberValue_serial_dates_match_ToInvariantString()
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
            Assert.AreEqual(serial.ToInvariantString(), WriteDouble(serial), $"date #{i}: {date:O}");
        }
    }

    [TestCase(0)]
    [TestCase(1)]
    [TestCase(-1)]
    [TestCase(12345)]
    [TestCase(int.MaxValue)]
    [TestCase(int.MinValue)]
    public void WriteNumberValue_int_matches_ToInvariantString(int value)
    {
        Assert.AreEqual(value.ToInvariantString(), WriteInt(value));
    }

    [TestCase(0u)]
    [TestCase(1u)]
    [TestCase(12345u)]
    [TestCase(uint.MaxValue)]
    public void WriteNumberValue_uint_matches_ToInvariantString(uint value)
    {
        Assert.AreEqual(value.ToInvariantString(), WriteUInt(value));
    }

    [Test]
    public void WriteNumberValue_int_matches_ToInvariantString_for_random_values()
    {
#pragma warning disable S2245 // Deterministic seed keeps failures reproducible
        var random = new Random(99);
#pragma warning restore S2245

        for (var i = 0; i < 20_000; i++)
        {
            var value = random.Next(int.MinValue, int.MaxValue);
            Assert.AreEqual(value.ToInvariantString(), WriteInt(value), $"value #{i}");
        }
    }

    /// <summary>
    /// The number buffer is thread-static and reused across calls, so a second value must not be
    /// contaminated by the leftovers of a longer first one.
    /// </summary>
    [Test]
    public void WriteNumberValue_reuses_the_buffer_without_leaking_previous_digits()
    {
        // "G15", not round-trip: 15 significant digits, matching ToInvariantString.
        Assert.AreEqual("-1.79769313486232E+308", WriteDouble(double.MinValue));
        Assert.AreEqual("1", WriteDouble(1d));
        Assert.AreEqual("-2147483648", WriteInt(int.MinValue));
        Assert.AreEqual("7", WriteInt(7));
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
