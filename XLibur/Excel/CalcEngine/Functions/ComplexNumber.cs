using System;
using System.Globalization;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// A complex number as Excel's IM* functions see it. Excel has no complex value type — a complex
/// number is the <em>text</em> "3+4i", so every one of those functions parses its arguments out of
/// text and writes its result back as text. This type carries the imaginary unit suffix along with
/// the components, because Excel echoes back whichever of "i" or "j" the input used.
/// </summary>
internal readonly record struct ComplexNumber(double Real, double Imaginary, char Suffix)
{
    /// <summary>The suffix Excel uses when nothing in the input said otherwise.</summary>
    internal const char DefaultSuffix = 'i';

    internal static ComplexNumber Zero => new(0, 0, DefaultSuffix);

    internal double Modulus => Math.Sqrt(Real * Real + Imaginary * Imaginary);

    /// <summary>The argument θ in (-π, π], undefined at the origin.</summary>
    internal double Argument => Math.Atan2(Imaginary, Real);

    internal bool IsZero => Real == 0 && Imaginary == 0;

    internal ComplexNumber WithSuffix(char suffix) => new(Real, Imaginary, suffix);

    internal ComplexNumber Conjugate() => new(Real, -Imaginary, Suffix);

    /// <summary>The complex number 1 carrying this one's suffix — the numerator of sec, csc, sech and csch.</summary>
    internal ComplexNumber One() => new(1, 0, Suffix);

    internal ComplexNumber Add(in ComplexNumber other) => new(Real + other.Real, Imaginary + other.Imaginary, Suffix);

    internal ComplexNumber Subtract(in ComplexNumber other) => new(Real - other.Real, Imaginary - other.Imaginary, Suffix);

    internal ComplexNumber Multiply(in ComplexNumber other)
        => new(Real * other.Real - Imaginary * other.Imaginary,
               Real * other.Imaginary + Imaginary * other.Real,
               Suffix);

    internal ComplexNumber Divide(in ComplexNumber other)
    {
        var denominator = other.Real * other.Real + other.Imaginary * other.Imaginary;
        return new((Real * other.Real + Imaginary * other.Imaginary) / denominator,
                   (Imaginary * other.Real - Real * other.Imaginary) / denominator,
                   Suffix);
    }

    /// <summary>e^z = e^re · (cos im + i·sin im).</summary>
    internal ComplexNumber Exp()
    {
        var magnitude = Math.Exp(Real);
        return new(magnitude * Math.Cos(Imaginary), magnitude * Math.Sin(Imaginary), Suffix);
    }

    /// <summary>The principal natural logarithm: ln|z| + i·arg z.</summary>
    internal ComplexNumber Log() => new(Math.Log(Modulus), Argument, Suffix);

    /// <summary>z^n for real n, via the polar form.</summary>
    internal ComplexNumber Power(double exponent)
    {
        var magnitude = Math.Pow(Modulus, exponent);
        var angle = Argument * exponent;
        return new(magnitude * Math.Cos(angle), magnitude * Math.Sin(angle), Suffix);
    }

    internal ComplexNumber Sqrt() => Power(0.5);

    internal ComplexNumber Sin()
        => new(Math.Sin(Real) * Math.Cosh(Imaginary), Math.Cos(Real) * Math.Sinh(Imaginary), Suffix);

    internal ComplexNumber Cos()
        => new(Math.Cos(Real) * Math.Cosh(Imaginary), -Math.Sin(Real) * Math.Sinh(Imaginary), Suffix);

    internal ComplexNumber Sinh()
        => new(Math.Sinh(Real) * Math.Cos(Imaginary), Math.Cosh(Real) * Math.Sin(Imaginary), Suffix);

    internal ComplexNumber Cosh()
        => new(Math.Cosh(Real) * Math.Cos(Imaginary), Math.Sinh(Real) * Math.Sin(Imaginary), Suffix);

    /// <summary>
    /// Parse Excel's complex-number text: an optional real part, an optional signed imaginary part
    /// and the "i" or "j" that marks it. "3+4i", "3", "4i", "-i" and "1.5E+3-2j" are all valid; the
    /// suffix is case sensitive, as it is in Excel.
    /// </summary>
    internal static bool TryParse(string text, out ComplexNumber value)
    {
        value = Zero;
        var span = text.AsSpan().Trim();
        if (span.IsEmpty)
        {
            // Excel reads an empty argument as zero rather than rejecting it.
            return true;
        }

        var suffix = span[^1];
        if (suffix is not ('i' or 'j'))
        {
            // No suffix at all: the whole thing has to be a plain real number.
            if (!TryParseComponent(span, out var realOnly))
                return false;

            value = new ComplexNumber(realOnly, 0, DefaultSuffix);
            return true;
        }

        var body = span[..^1];
        var split = FindImaginarySignIndex(body);

        ReadOnlySpan<char> realPart;
        ReadOnlySpan<char> imaginaryPart;
        if (split < 0)
        {
            realPart = default;
            imaginaryPart = body;
        }
        else
        {
            realPart = body[..split];
            imaginaryPart = body[split..];
        }

        double real = 0;
        if (!realPart.IsEmpty && !TryParseComponent(realPart, out real))
            return false;

        if (!TryParseImaginaryComponent(imaginaryPart, out var imaginary))
            return false;

        value = new ComplexNumber(real, imaginary, suffix);
        return true;
    }

    /// <summary>
    /// Index of the sign that separates the real part from the imaginary one, or -1 when the text
    /// holds only an imaginary part. A sign at the very start belongs to the real part, and one
    /// straight after an exponent marker belongs to that exponent.
    /// </summary>
    private static int FindImaginarySignIndex(ReadOnlySpan<char> body)
    {
        for (var i = body.Length - 1; i > 0; i--)
        {
            if (body[i] is not ('+' or '-'))
                continue;

            if (body[i - 1] is 'e' or 'E')
                continue;

            return i;
        }

        return -1;
    }

    /// <summary>
    /// Parse the imaginary coefficient, which Excel lets you leave out entirely: "i" is 1 and "-i"
    /// is -1.
    /// </summary>
    private static bool TryParseImaginaryComponent(ReadOnlySpan<char> text, out double value)
    {
        if (text.IsEmpty || text.SequenceEqual("+".AsSpan()))
        {
            value = 1;
            return true;
        }

        if (text.SequenceEqual("-".AsSpan()))
        {
            value = -1;
            return true;
        }

        return TryParseComponent(text, out value);
    }

    private static bool TryParseComponent(ReadOnlySpan<char> text, out double value)
    {
        // Complex-number text is always written with a period, whatever the workbook's culture.
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
               && !double.IsNaN(value)
               && !double.IsInfinity(value);
    }

    /// <summary>
    /// Render back to Excel's text form: a purely real number drops the imaginary part, a purely
    /// imaginary one drops the real part, and a unit coefficient is written as bare "i".
    /// </summary>
    public override string ToString()
    {
        var suffix = Suffix == default ? DefaultSuffix : Suffix;

        if (Imaginary == 0)
            return FormatComponent(Real);

        var imaginary = FormatImaginary(Imaginary, suffix);
        if (Real == 0)
            return imaginary;

        // FormatImaginary already carries the sign of a negative coefficient.
        var separator = Imaginary > 0 ? "+" : string.Empty;
        return FormatComponent(Real) + separator + imaginary;
    }

    private static string FormatImaginary(double imaginary, char suffix)
    {
        if (imaginary == 1)
            return suffix.ToString();

        if (imaginary == -1)
            return "-" + suffix;

        return FormatComponent(imaginary) + suffix;
    }

    /// <summary>
    /// Excel writes the components of a complex number with 15 significant digits, which is also
    /// what hides the last-bit noise of the trigonometric identities — IMSQRT("3+4i") comes out as
    /// "2+i" rather than "2.00000000000000018+i".
    /// </summary>
    private static string FormatComponent(double value)
        => value.ToString("G15", CultureInfo.InvariantCulture);
}
