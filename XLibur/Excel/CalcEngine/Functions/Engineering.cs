using System;
using System.Collections.Generic;
using XLibur.Excel.CalcEngine.Functions;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Excel.CalcEngine;

internal static class Engineering
{
    // Maximum values for each base (10 characters each, two's complement)
    // BIN: 10 bits, range -512 to 511
    // OCT: 10 digits = 30 bits, range -536870912 to 536870911
    // HEX: 10 digits = 40 bits, range -549755813888 to 549755813887
    private const long BinMax = 511;
    private const long BinMin = -512;
    private const long OctMax = 536870911;
    private const long OctMin = -536870912;
    private const long HexMax = 549755813887;
    private const long HexMin = -549755813888;

    /// <summary>Excel's bitwise functions operate on 48-bit unsigned integers.</summary>
    private const double BitMax = 281474976710655; // 2^48 - 1

    /// <summary>A shift may not move a bit past the largest integer a double represents exactly.</summary>
    private const double ShiftedMax = 9007199254740991; // 2^53 - 1

    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("BESSELI", 2, 2, Adapt(BesselI), FunctionFlags.Scalar); // Returns the modified Bessel function In(x)
        ce.RegisterFunction("BESSELJ", 2, 2, Adapt(BesselJ), FunctionFlags.Scalar); // Returns the Bessel function Jn(x)
        ce.RegisterFunction("BESSELK", 2, 2, Adapt(BesselK), FunctionFlags.Scalar); // Returns the modified Bessel function Kn(x)
        ce.RegisterFunction("BESSELY", 2, 2, Adapt(BesselY), FunctionFlags.Scalar); // Returns the Bessel function Yn(x)
        ce.RegisterFunction("BIN2DEC", 1, 1, Adapt(Bin2Dec), FunctionFlags.Scalar);
        ce.RegisterFunction("BIN2HEX", 1, 2, AdaptLastOptional(Bin2Hex, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("BIN2OCT", 1, 2, AdaptLastOptional(Bin2Oct, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("BITAND", 2, 2, Adapt(BitAnd), FunctionFlags.Scalar | FunctionFlags.Future); // Returns a bitwise 'And' of two numbers
        ce.RegisterFunction("BITLSHIFT", 2, 2, Adapt(BitLShift), FunctionFlags.Scalar | FunctionFlags.Future); // Returns a number shifted left by shift_amount bits
        ce.RegisterFunction("BITOR", 2, 2, Adapt(BitOr), FunctionFlags.Scalar | FunctionFlags.Future); // Returns a bitwise 'Or' of two numbers
        ce.RegisterFunction("BITRSHIFT", 2, 2, Adapt(BitRShift), FunctionFlags.Scalar | FunctionFlags.Future); // Returns a number shifted right by shift_amount bits
        ce.RegisterFunction("BITXOR", 2, 2, Adapt(BitXor), FunctionFlags.Scalar | FunctionFlags.Future); // Returns a bitwise 'Exclusive Or' of two numbers
        ce.RegisterFunction("COMPLEX", 2, 3, AdaptLastOptional(Complex, "i"), FunctionFlags.Scalar); // Converts real and imaginary coefficients into a complex number
        ce.RegisterFunction("CONVERT", 3, 3, Adapt(ConvertUnit), FunctionFlags.Scalar); // Converts a number from one measurement system to another
        ce.RegisterFunction("DEC2BIN", 1, 2, AdaptLastOptional(Dec2Bin, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("DEC2HEX", 1, 2, AdaptLastOptional(Dec2Hex, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("DEC2OCT", 1, 2, AdaptLastOptional(Dec2Oct, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("DELTA", 1, 2, AdaptLastOptional(Delta, 0), FunctionFlags.Scalar); // Tests whether two values are equal
        ce.RegisterFunction("ERF", 1, 2, AdaptLastOptional(Erf), FunctionFlags.Scalar); // Returns the error function
        ce.RegisterFunction("ERF.PRECISE", 1, 1, Adapt(ErfPrecise), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the error function
        ce.RegisterFunction("ERFC", 1, 1, Adapt(Erfc), FunctionFlags.Scalar); // Returns the complementary error function
        ce.RegisterFunction("ERFC.PRECISE", 1, 1, Adapt(Erfc), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the complementary ERF function integrated between x and infinity
        ce.RegisterFunction("GESTEP", 1, 2, AdaptLastOptional(GeStep, 0), FunctionFlags.Scalar); // Tests whether a number is greater than a threshold value
        ce.RegisterFunction("HEX2BIN", 1, 2, AdaptLastOptional(Hex2Bin, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("HEX2DEC", 1, 1, Adapt(Hex2Dec), FunctionFlags.Scalar);
        ce.RegisterFunction("HEX2OCT", 1, 2, AdaptLastOptional(Hex2Oct, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("IMABS", 1, 1, Adapt(ImAbs), FunctionFlags.Scalar); // Returns the absolute value(modulus) of a complex number
        ce.RegisterFunction("IMAGINARY", 1, 1, Adapt(ImAginary), FunctionFlags.Scalar); // Returns the imaginary coefficient of a complex number
        ce.RegisterFunction("IMARGUMENT", 1, 1, Adapt(ImArgument), FunctionFlags.Scalar); // Returns the argument theta, an angle expressed in radians
        ce.RegisterFunction("IMCONJUGATE", 1, 1, Adapt(ImConjugate), FunctionFlags.Scalar); // Returns the complex conjugate of a complex number
        ce.RegisterFunction("IMCOS", 1, 1, Adapt(ImCos), FunctionFlags.Scalar); // Returns the cosine of a complex number
        ce.RegisterFunction("IMCOSH", 1, 1, Adapt(ImCosh), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the hyperbolic cosine of a complex number
        ce.RegisterFunction("IMCOT", 1, 1, Adapt(ImCot), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the cotangent of a complex number
        ce.RegisterFunction("IMCSC", 1, 1, Adapt(ImCsc), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the cosecant of a complex number
        ce.RegisterFunction("IMCSCH", 1, 1, Adapt(ImCsch), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the hyperbolic cosecant of a complex number
        ce.RegisterFunction("IMDIV", 2, 2, Adapt(ImDiv), FunctionFlags.Scalar); // Returns the quotient of two complex numbers
        ce.RegisterFunction("IMEXP", 1, 1, Adapt(ImExp), FunctionFlags.Scalar); // Returns the exponential of a complex number
        ce.RegisterFunction("IMLN", 1, 1, Adapt(ImLn), FunctionFlags.Scalar); // Returns the natural logarithm of a complex number
        ce.RegisterFunction("IMLOG10", 1, 1, Adapt(ImLog10), FunctionFlags.Scalar); // Returns the base - 10 logarithm of a complex number
        ce.RegisterFunction("IMLOG2", 1, 1, Adapt(ImLog2), FunctionFlags.Scalar); // Returns the base - 2 logarithm of a complex number
        ce.RegisterFunction("IMPOWER", 2, 2, Adapt(ImPower), FunctionFlags.Scalar); // Returns a complex number raised to an integer power
        ce.RegisterFunction("IMPRODUCT", 1, 255, Adapt(ImProduct), FunctionFlags.Scalar); // Returns the product of from 2 to 255 complex numbers
        ce.RegisterFunction("IMREAL", 1, 1, Adapt(ImReal), FunctionFlags.Scalar); // Returns the real coefficient of a complex number
        ce.RegisterFunction("IMSEC", 1, 1, Adapt(ImSec), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the secant of a complex number
        ce.RegisterFunction("IMSECH", 1, 1, Adapt(ImSech), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the hyperbolic secant of a complex number
        ce.RegisterFunction("IMSIN", 1, 1, Adapt(ImSin), FunctionFlags.Scalar); // Returns the sine of a complex number
        ce.RegisterFunction("IMSINH", 1, 1, Adapt(ImSinh), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the hyperbolic sine of a complex number
        ce.RegisterFunction("IMSQRT", 1, 1, Adapt(ImSqrt), FunctionFlags.Scalar); // Returns the square root of a complex number
        ce.RegisterFunction("IMSUB", 2, 2, Adapt(ImSub), FunctionFlags.Scalar); // Returns the difference between two complex numbers
        ce.RegisterFunction("IMSUM", 1, 255, Adapt(ImSum), FunctionFlags.Scalar); // Returns the sum of complex numbers
        ce.RegisterFunction("IMTAN", 1, 1, Adapt(ImTan), FunctionFlags.Scalar | FunctionFlags.Future); // Returns the tangent of a complex number
        ce.RegisterFunction("OCT2BIN", 1, 2, AdaptLastOptional(Oct2Bin, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("OCT2DEC", 1, 1, Adapt(Oct2Dec), FunctionFlags.Scalar);
        ce.RegisterFunction("OCT2HEX", 1, 2, AdaptLastOptional(Oct2Hex, 0), FunctionFlags.Scalar);
    }

    #region Bessel functions

    private static ScalarValue BesselI(CalcContext ctx, double x, double order)
        => TryGetBesselOrder(order, out var n) ? Bessel.I(x, n) : XLError.NumberInvalid;

    private static ScalarValue BesselJ(CalcContext ctx, double x, double order)
        => TryGetBesselOrder(order, out var n) ? Bessel.J(x, n) : XLError.NumberInvalid;

    private static ScalarValue BesselK(CalcContext ctx, double x, double order)
    {
        // K is singular at the origin and complex for negative arguments.
        if (x <= 0 || !TryGetBesselOrder(order, out var n))
            return XLError.NumberInvalid;

        return Bessel.K(x, n);
    }

    private static ScalarValue BesselY(CalcContext ctx, double x, double order)
    {
        if (x <= 0 || !TryGetBesselOrder(order, out var n))
            return XLError.NumberInvalid;

        return Bessel.Y(x, n);
    }

    private static bool TryGetBesselOrder(double order, out int n)
    {
        n = 0;
        var truncated = Math.Truncate(order);
        if (truncated < 0 || truncated > int.MaxValue)
            return false;

        n = (int)truncated;
        return true;
    }

    #endregion

    #region Bitwise functions

    private static ScalarValue BitAnd(CalcContext ctx, double number1, double number2)
        => Bitwise(number1, number2, static (a, b) => a & b);

    private static ScalarValue BitOr(CalcContext ctx, double number1, double number2)
        => Bitwise(number1, number2, static (a, b) => a | b);

    private static ScalarValue BitXor(CalcContext ctx, double number1, double number2)
        => Bitwise(number1, number2, static (a, b) => a ^ b);

    private static ScalarValue Bitwise(double number1, double number2, Func<long, long, long> combine)
    {
        if (!TryGetBitOperand(number1, out var a) || !TryGetBitOperand(number2, out var b))
            return XLError.NumberInvalid;

        return (double)combine(a, b);
    }

    private static ScalarValue BitLShift(CalcContext ctx, double number, double shift)
        => Shift(number, shift);

    private static ScalarValue BitRShift(CalcContext ctx, double number, double shift)
        => Shift(number, -shift);

    /// <summary>
    /// Shift left by <paramref name="shift"/> bits, or right when it is negative. Excel expresses
    /// both directions as one operation, so BITRSHIFT is BITLSHIFT with the sign flipped.
    /// </summary>
    private static ScalarValue Shift(double number, double shift)
    {
        if (!TryGetBitOperand(number, out var value))
            return XLError.NumberInvalid;

        var amount = Math.Truncate(shift);
        if (Math.Abs(amount) > 53)
            return XLError.NumberInvalid;

        // Shifting is defined arithmetically rather than on the bit pattern, which keeps the result
        // exact for the 48-bit inputs and lets a shift beyond 53 bits be caught as out of range.
        var result = value * Math.Pow(2, amount);
        result = amount < 0 ? Math.Floor(result) : result;
        if (result > ShiftedMax)
            return XLError.NumberInvalid;

        return result;
    }

    private static bool TryGetBitOperand(double number, out long value)
    {
        value = 0;
        if (number < 0 || number > BitMax || number != Math.Truncate(number))
            return false;

        value = (long)number;
        return true;
    }

    #endregion

    #region Comparison and error functions

    private static ScalarValue Delta(CalcContext ctx, double number1, double number2)
        => number1 == number2 ? 1d : 0d;

    private static ScalarValue GeStep(CalcContext ctx, double number, double step)
        => number >= step ? 1d : 0d;

    /// <summary>
    /// ERF(lower, [upper]) integrates the error function between the two limits; with one argument
    /// it integrates from zero, which is erf(lower).
    /// </summary>
    private static ScalarValue Erf(CalcContext ctx, double lowerLimit, double? upperLimit)
    {
        if (upperLimit is null)
            return XLMath.Erf(lowerLimit);

        return XLMath.Erf(upperLimit.Value) - XLMath.Erf(lowerLimit);
    }

    private static ScalarValue ErfPrecise(CalcContext ctx, double x) => XLMath.Erf(x);

    private static ScalarValue Erfc(CalcContext ctx, double x) => XLMath.Erfc(x);

    #endregion

    #region Unit conversion

    private static ScalarValue ConvertUnit(CalcContext ctx, double number, string fromUnit, string toUnit)
    {
        // An unknown unit, or two units that measure different things, is #N/A rather than #VALUE!:
        // the arguments were the right type, there is just no conversion between them.
        if (!UnitConversion.TryConvert(number, fromUnit, toUnit, out var result))
            return XLError.NoValueAvailable;

        return result;
    }

    #endregion

    #region Complex numbers

    private static ScalarValue Complex(CalcContext ctx, double real, double imaginary, string suffix)
    {
        if (suffix is not ("i" or "j"))
            return XLError.IncompatibleValue;

        return new ComplexNumber(real, imaginary, suffix[0]).ToString();
    }

    private static ScalarValue ImAbs(CalcContext ctx, string number)
        => WithComplex(number, static z => z.Modulus);

    private static ScalarValue ImReal(CalcContext ctx, string number)
        => WithComplex(number, static z => z.Real);

    private static ScalarValue ImAginary(CalcContext ctx, string number)
        => WithComplex(number, static z => z.Imaginary);

    private static ScalarValue ImArgument(CalcContext ctx, string number)
    {
        if (!ComplexNumber.TryParse(number, out var z))
            return XLError.NumberInvalid;

        // The argument of zero is not defined; Excel reports that as a division by zero.
        if (z.IsZero)
            return XLError.DivisionByZero;

        return z.Argument;
    }

    private static ScalarValue ImConjugate(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Conjugate());

    private static ScalarValue ImExp(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Exp());

    private static ScalarValue ImLn(CalcContext ctx, string number)
        => MapDefinedComplex(number, static z => z.Log());

    private static ScalarValue ImLog10(CalcContext ctx, string number)
        => MapDefinedComplex(number, static z => Scale(z.Log(), 1 / Math.Log(10)));

    private static ScalarValue ImLog2(CalcContext ctx, string number)
        => MapDefinedComplex(number, static z => Scale(z.Log(), 1 / Math.Log(2)));

    private static ScalarValue ImSqrt(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Sqrt());

    private static ScalarValue ImSin(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Sin());

    private static ScalarValue ImCos(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Cos());

    private static ScalarValue ImSinh(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Sinh());

    private static ScalarValue ImCosh(CalcContext ctx, string number)
        => MapComplex(number, static z => z.Cosh());

    private static ScalarValue ImTan(CalcContext ctx, string number)
        => MapRatio(number, static z => z.Sin(), static z => z.Cos());

    private static ScalarValue ImCot(CalcContext ctx, string number)
        => MapRatio(number, static z => z.Cos(), static z => z.Sin());

    private static ScalarValue ImSec(CalcContext ctx, string number)
        => MapRatio(number, static z => z.One(), static z => z.Cos());

    private static ScalarValue ImCsc(CalcContext ctx, string number)
        => MapRatio(number, static z => z.One(), static z => z.Sin());

    private static ScalarValue ImSech(CalcContext ctx, string number)
        => MapRatio(number, static z => z.One(), static z => z.Cosh());

    private static ScalarValue ImCsch(CalcContext ctx, string number)
        => MapRatio(number, static z => z.One(), static z => z.Sinh());

    private static ScalarValue ImPower(string number, double power)
    {
        if (!ComplexNumber.TryParse(number, out var z))
            return XLError.NumberInvalid;

        // 0^0 and 0 to a negative power are both undefined.
        if (z.IsZero && power <= 0)
            return XLError.NumberInvalid;

        return z.Power(power).ToString();
    }

    private static ScalarValue ImSub(string number1, string number2)
        => Combine(number1, number2, static (a, b) => a.Subtract(b), divide: false);

    private static ScalarValue ImDiv(string number1, string number2)
        => Combine(number1, number2, static (a, b) => a.Divide(b), divide: true);

    private static ScalarValue ImSum(CalcContext ctx, List<string> numbers)
        => Accumulate(numbers, ComplexNumber.Zero, static (a, b) => a.Add(b));

    private static ScalarValue ImProduct(CalcContext ctx, List<string> numbers)
        => Accumulate(numbers, new ComplexNumber(1, 0, ComplexNumber.DefaultSuffix), static (a, b) => a.Multiply(b));

    private static ScalarValue WithComplex(string number, Func<ComplexNumber, double> selector)
    {
        if (!ComplexNumber.TryParse(number, out var z))
            return XLError.NumberInvalid;

        return selector(z);
    }

    private static ScalarValue MapComplex(string number, Func<ComplexNumber, ComplexNumber> selector)
    {
        if (!ComplexNumber.TryParse(number, out var z))
            return XLError.NumberInvalid;

        return selector(z).ToString();
    }

    /// <summary>Like <see cref="MapComplex"/>, for the functions that are singular at the origin.</summary>
    private static ScalarValue MapDefinedComplex(string number, Func<ComplexNumber, ComplexNumber> selector)
    {
        if (!ComplexNumber.TryParse(number, out var z))
            return XLError.NumberInvalid;

        if (z.IsZero)
            return XLError.NumberInvalid;

        return selector(z).ToString();
    }

    /// <summary>
    /// The reciprocal-style functions (tan, cot, sec, csc and their hyperbolic partners) are all a
    /// quotient of two of the primitives, and all report a zero denominator as #NUM!.
    /// </summary>
    private static ScalarValue MapRatio(string number, Func<ComplexNumber, ComplexNumber> numerator, Func<ComplexNumber, ComplexNumber> denominator)
    {
        if (!ComplexNumber.TryParse(number, out var z))
            return XLError.NumberInvalid;

        var bottom = denominator(z);
        if (bottom.IsZero)
            return XLError.NumberInvalid;

        return numerator(z).Divide(bottom).ToString();
    }

    private static ScalarValue Combine(string number1, string number2, Func<ComplexNumber, ComplexNumber, ComplexNumber> combine, bool divide)
    {
        if (!ComplexNumber.TryParse(number1, out var a) || !ComplexNumber.TryParse(number2, out var b))
            return XLError.NumberInvalid;

        if (!TryAgreeOnSuffix(a, b, out var suffix))
            return XLError.IncompatibleValue;

        if (divide && b.IsZero)
            return XLError.NumberInvalid;

        return combine(a.WithSuffix(suffix), b).ToString();
    }

    private static ScalarValue Accumulate(List<string> numbers, ComplexNumber seed, Func<ComplexNumber, ComplexNumber, ComplexNumber> combine)
    {
        var total = seed;
        var suffix = default(char);
        foreach (var text in numbers)
        {
            if (!ComplexNumber.TryParse(text, out var z))
                return XLError.NumberInvalid;

            if (!TryTakeSuffix(ref suffix, z))
                return XLError.IncompatibleValue;

            total = combine(total, z);
        }

        return total.WithSuffix(suffix == default ? ComplexNumber.DefaultSuffix : suffix).ToString();
    }

    /// <summary>
    /// Excel refuses to mix the two spellings of the imaginary unit: IMSUM("1+i", "1+j") is
    /// #VALUE!. A number with no imaginary part carries no suffix and agrees with anything.
    /// </summary>
    private static bool TryAgreeOnSuffix(in ComplexNumber a, in ComplexNumber b, out char suffix)
    {
        suffix = ComplexNumber.DefaultSuffix;
        var first = a.Imaginary == 0 ? default : a.Suffix;
        var second = b.Imaginary == 0 ? default : b.Suffix;

        if (first != default && second != default && first != second)
            return false;

        if (first != default)
            suffix = first;
        else if (second != default)
            suffix = second;

        return true;
    }

    private static bool TryTakeSuffix(ref char suffix, in ComplexNumber z)
    {
        if (z.Imaginary == 0)
            return true;

        if (suffix == default)
        {
            suffix = z.Suffix;
            return true;
        }

        return suffix == z.Suffix;
    }

    private static ComplexNumber Scale(in ComplexNumber z, double factor)
        => new(z.Real * factor, z.Imaginary * factor, z.Suffix);

    #endregion

    /// <summary>
    /// Parse a string as a number in the given base. Returns the signed value using two's complement
    /// with the specified bit width.
    /// </summary>
    private static bool TryParseBase(string text, int fromBase, int bitWidth, out long value)
    {
        value = 0;
        text = text.Trim();
        if (text.Length == 0 || text.Length > 10)
            return false;

        try
        {
            value = Convert.ToInt64(text, fromBase);
        }
        catch (FormatException)
        {
            return false;
        }
        catch (OverflowException)
        {
            return false;
        }

        // Apply two's complement for negative numbers.
        // If the highest bit is set, the value is negative.
        var maxUnsigned = 1L << bitWidth;
        if (value >= maxUnsigned)
            return false;

        var signBit = 1L << (bitWidth - 1);
        if ((value & signBit) != 0)
            value -= maxUnsigned;

        return true;
    }

    /// <summary>
    /// Convert a signed value to a string in the given base using two's complement with specified bit width.
    /// </summary>
    private static string ToBaseString(long value, int toBase, int bitWidth)
    {
        if (value < 0)
        {
            // Two's complement: add 2^bitWidth to get the unsigned representation
            value += 1L << bitWidth;
        }

        var result = Convert.ToString(value, toBase).ToUpperInvariant();
        return result;
    }

    private static ScalarValue ApplyPlaces(string result, double placesDouble)
    {
        var places = (int)Math.Truncate(placesDouble);
        if (places == 0)
            return result;

        if (places < 0 || places > 10)
            return XLError.NumberInvalid;

        // Places only applies to non-negative results (no leading F's/7's/1's for padding)
        if (result.Length > places)
            return XLError.NumberInvalid;

        return result.PadLeft(places, '0');
    }

    #region BIN2*

    private static ScalarValue Bin2Dec(CalcContext ctx, string number)
    {
        if (!TryParseBase(number, 2, 10, out var value))
            return XLError.NumberInvalid;

        return (double)value;
    }

    private static ScalarValue Bin2Hex(CalcContext ctx, string number, double places)
    {
        if (!TryParseBase(number, 2, 10, out var value))
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 16, 40);
        return ApplyPlaces(result, places);
    }

    private static ScalarValue Bin2Oct(CalcContext ctx, string number, double places)
    {
        if (!TryParseBase(number, 2, 10, out var value))
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 8, 30);
        return ApplyPlaces(result, places);
    }

    #endregion

    #region DEC2*

    private static ScalarValue Dec2Bin(CalcContext ctx, double number, double places)
    {
        var value = (long)Math.Truncate(number);
        if (value < BinMin || value > BinMax)
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 2, 10);
        return ApplyPlaces(result, places);
    }

    private static ScalarValue Dec2Hex(CalcContext ctx, double number, double places)
    {
        var value = (long)Math.Truncate(number);
        if (value < HexMin || value > HexMax)
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 16, 40);
        return ApplyPlaces(result, places);
    }

    private static ScalarValue Dec2Oct(CalcContext ctx, double number, double places)
    {
        var value = (long)Math.Truncate(number);
        if (value < OctMin || value > OctMax)
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 8, 30);
        return ApplyPlaces(result, places);
    }

    #endregion

    #region HEX2*

    private static ScalarValue Hex2Bin(CalcContext ctx, string number, double places)
    {
        if (!TryParseBase(number, 16, 40, out var value))
            return XLError.NumberInvalid;

        if (value < BinMin || value > BinMax)
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 2, 10);
        return ApplyPlaces(result, places);
    }

    private static ScalarValue Hex2Dec(CalcContext ctx, string number)
    {
        if (!TryParseBase(number, 16, 40, out var value))
            return XLError.NumberInvalid;

        return (double)value;
    }

    private static ScalarValue Hex2Oct(CalcContext ctx, string number, double places)
    {
        if (!TryParseBase(number, 16, 40, out var value))
            return XLError.NumberInvalid;

        if (value < OctMin || value > OctMax)
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 8, 30);
        return ApplyPlaces(result, places);
    }

    #endregion

    #region OCT2*

    private static ScalarValue Oct2Bin(CalcContext ctx, string number, double places)
    {
        if (!TryParseBase(number, 8, 30, out var value))
            return XLError.NumberInvalid;

        if (value < BinMin || value > BinMax)
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 2, 10);
        return ApplyPlaces(result, places);
    }

    private static ScalarValue Oct2Dec(CalcContext ctx, string number)
    {
        if (!TryParseBase(number, 8, 30, out var value))
            return XLError.NumberInvalid;

        return (double)value;
    }

    private static ScalarValue Oct2Hex(CalcContext ctx, string number, double places)
    {
        if (!TryParseBase(number, 8, 30, out var value))
            return XLError.NumberInvalid;

        var result = ToBaseString(value, 16, 40);
        return ApplyPlaces(result, places);
    }

    #endregion
}
