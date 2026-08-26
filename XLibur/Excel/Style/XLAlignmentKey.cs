using System;

namespace XLibur.Excel;

internal readonly record struct XLAlignmentKey
{
    private readonly XLAlignmentHorizontalValues _horizontal;
    private readonly XLAlignmentVerticalValues _vertical;
    private readonly XLAlignmentReadingOrderValues _readingOrder;

    /// <summary>
    /// Rejects an enum value that is not one of its type's named members, rather than storing it
    /// unexamined.
    /// </summary>
    /// <remarks>
    /// <see cref="XLAlignmentHorizontalValues"/>, <see cref="XLAlignmentVerticalValues"/> and
    /// <see cref="XLAlignmentReadingOrderValues"/> are public and already declared with a
    /// <c>byte</c> underlying type, unlike <c>XLBorderStyleValues</c>, <c>XLColorType</c> and
    /// <c>XLThemeColor</c>, which are <c>int</c>-backed and narrowed into a byte field by
    /// <c>XLBorderKey</c> and <c>XLColorKey</c>. That difference means the numeric-range check
    /// those two use - "does the value fit in a byte" - would do nothing useful here: a value of a
    /// byte-backed enum always fits in a byte by construction, so the check would never reject
    /// anything. What is still possible, because C# permits an explicit cast from any numeric type
    /// to any enum regardless of its underlying width or which members are defined, is a value with
    /// no defined member at all - <c>(XLAlignmentHorizontalValues)999</c> compiles and, cast
    /// unchecked to fit the enum's byte storage, silently becomes
    /// <c>(XLAlignmentHorizontalValues)231</c>. Checking membership is the only guard against that
    /// left available once the storage width itself can no longer be the check.
    /// <para>
    /// The two guards are not equally complete, and this one cannot be made to match the other:
    /// <c>XLBorderKey</c>'s narrowing runs inside this library's own code, on a value still in its
    /// original, wider <c>int</c> form, so it can reject an out-of-range value before any wrap
    /// happens - the aliasing it protects against never occurs. Here, the narrowing that can wrap a
    /// value onto a real member happens inside the caller's own cast expression, entirely before any
    /// code in this library runs, so a value that lands on a defined member - <c>262</c> wraps to
    /// <c>6</c>, which is <see cref="XLAlignmentHorizontalValues.Left"/> - arrives indistinguishable
    /// from a caller who passed <c>Left</c> outright. No check placed anywhere in this library can
    /// tell the two apart, because by the time either reaches here, they are the same value. What
    /// membership checking still closes is the much larger remainder: of a byte's 256 possible
    /// values, only 8 are members of this enum, so most out-of-range casts land somewhere with no
    /// name at all, and those this catches.
    /// </para>
    /// <para>
    /// Applied once, here, rather than on the <c>XLAlignment</c> facade that exposes these
    /// properties to a caller: every path that builds
    /// or modifies a key, including <c>StyleDecoder.AlignmentKey</c> and the
    /// external-<c>IXLAlignment</c> branch of <c>XLAlignment.GenerateKey</c>, goes through one of
    /// these three <c>init</c> accessors or a <c>with</c> expression backed by them, so a future
    /// caller cannot bypass the check by reaching the key through a path this file did not
    /// anticipate.
    /// </para>
    /// </remarks>
    private static TEnum Validate<TEnum>(TEnum value) where TEnum : struct, Enum
    {
        if (!Enum.IsDefined(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value), value,
                $"{typeof(TEnum).Name} has no member with the value {value:D}.");
        }

        return value;
    }

    public required XLAlignmentHorizontalValues Horizontal
    {
        get => _horizontal;
        init => _horizontal = Validate(value);
    }

    public required XLAlignmentVerticalValues Vertical
    {
        get => _vertical;
        init => _vertical = Validate(value);
    }

    public required int Indent { get; init; }

    public required bool JustifyLastLine { get; init; }

    public required XLAlignmentReadingOrderValues ReadingOrder
    {
        get => _readingOrder;
        init => _readingOrder = Validate(value);
    }

    public required int RelativeIndent { get; init; }

    public required bool ShrinkToFit { get; init; }

    public required int TextRotation { get; init; }

    public required bool WrapText { get; init; }

    public override string ToString()
    {
        return
            $"{Horizontal} {Vertical} {ReadingOrder} Indent: {Indent} RelativeIndent: {RelativeIndent} TextRotation: {TextRotation} " +
            (WrapText ? "WrapText" : "") +
            (JustifyLastLine ? "JustifyLastLine" : "");
    }
}
