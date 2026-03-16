using System;
using System.Globalization;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

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

    public static void Register(FunctionRegistry ce)
    {
        // BESSELI Returns the modified Bessel function In(x)
        // BESSELJ Returns the Bessel function Jn(x)
        // BESSELK Returns the modified Bessel function Kn(x)
        // BESSELY Returns the Bessel function Yn(x)
        ce.RegisterFunction("BIN2DEC", 1, 1, Adapt(Bin2Dec), FunctionFlags.Scalar);
        ce.RegisterFunction("BIN2HEX", 1, 2, AdaptLastOptional(Bin2Hex, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("BIN2OCT", 1, 2, AdaptLastOptional(Bin2Oct, 0), FunctionFlags.Scalar);
        // BITAND Returns a bitwise 'And' of two numbers
        // BITLSHIFT Returns a number shifted left by shift_amount bits
        // BITOR Returns a bitwise 'Or' of two numbers
        // BITRSHIFT Returns a number shifted right by shift_amount bits
        // BITXOR Returns a bitwise 'Exclusive Or' of two numbers
        // COMPLEX Converts real and imaginary coefficients into a complex number
        // CONVERT Converts a number from one measurement system to another
        ce.RegisterFunction("DEC2BIN", 1, 2, AdaptLastOptional(Dec2Bin, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("DEC2HEX", 1, 2, AdaptLastOptional(Dec2Hex, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("DEC2OCT", 1, 2, AdaptLastOptional(Dec2Oct, 0), FunctionFlags.Scalar);
        // DELTA Tests whether two values are equal
        // ERF Returns the error function
        // ERF.PRECISE Returns the error function
        // ERFC Returns the complementary error function
        // ERFC.PRECISE Returns the complementary ERF function integrated between x and infinity
        // GESTEP Tests whether a number is greater than a threshold value
        ce.RegisterFunction("HEX2BIN", 1, 2, AdaptLastOptional(Hex2Bin, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("HEX2DEC", 1, 1, Adapt(Hex2Dec), FunctionFlags.Scalar);
        ce.RegisterFunction("HEX2OCT", 1, 2, AdaptLastOptional(Hex2Oct, 0), FunctionFlags.Scalar);
        // IMABS Returns the absolute value(modulus) of a complex number
        // IMAGINARY Returns the imaginary coefficient of a complex number
        // IMARGUMENT Returns the argument theta, an angle expressed in radians
        // IMCONJUGATE Returns the complex conjugate of a complex number
        // IMCOS Returns the cosine of a complex number
        // IMCOSH Returns the hyperbolic cosine of a complex number
        // IMCOT Returns the cotangent of a complex number
        // IMCSC Returns the cosecant of a complex number
        // IMCSCH Returns the hyperbolic cosecant of a complex number
        // IMDIV Returns the quotient of two complex numbers
        // IMEXP Returns the exponential of a complex number
        // IMLN Returns the natural logarithm of a complex number
        // IMLOG10 Returns the base - 10 logarithm of a complex number
        // IMLOG2 Returns the base - 2 logarithm of a complex number
        // IMPOWER Returns a complex number raised to an integer power
        // IMPRODUCT Returns the product of from 2 to 255 complex numbers
        // IMREAL Returns the real coefficient of a complex number
        // IMSEC Returns the secant of a complex number
        // IMSECH Returns the hyperbolic secant of a complex number
        // IMSIN Returns the sine of a complex number
        // IMSINH Returns the hyperbolic sine of a complex number
        // IMSQRT Returns the square root of a complex number
        // IMSUB Returns the difference between two complex numbers
        // IMSUM Returns the sum of complex numbers
        // IMTAN Returns the tangent of a complex number
        ce.RegisterFunction("OCT2BIN", 1, 2, AdaptLastOptional(Oct2Bin, 0), FunctionFlags.Scalar);
        ce.RegisterFunction("OCT2DEC", 1, 1, Adapt(Oct2Dec), FunctionFlags.Scalar);
        ce.RegisterFunction("OCT2HEX", 1, 2, AdaptLastOptional(Oct2Hex, 0), FunctionFlags.Scalar);
    }

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
