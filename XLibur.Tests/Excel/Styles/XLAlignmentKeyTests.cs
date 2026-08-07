using System;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// <see cref="XLAlignmentHorizontalValues"/>, <see cref="XLAlignmentVerticalValues"/> and
/// <see cref="XLAlignmentReadingOrderValues"/> are public and already declared with a <c>byte</c>
/// underlying type. <c>XLAlignmentKey</c> stores them in explicit fields with a validating <c>init</c>
/// accessor rather than the plain auto-properties it used to have - see the type's remarks for why a
/// numeric-range check, the guard the <c>int</c>-backed enums <c>XLBorderKey</c> and <c>XLColorKey</c>
/// use, would not catch anything here, and for the residual gap the check that replaces it still
/// leaves - see <see cref="Wraparound_onto_a_defined_member_cannot_be_told_apart_from_that_member"/>.
/// </summary>
/// <remarks>
/// Introducing an explicit field alongside a custom accessor is exactly the change that would break
/// <c>XLAlignmentKey</c>'s equality and hash code silently if the compiler's synthesized
/// <c>record struct</c> members stopped seeing the field - they are not hand-written here, unlike
/// <c>XLBorderKey</c>'s. <see cref="Two_keys_equal_only_when_every_field_matches"/> is the check that
/// this still holds.
/// </remarks>
public class XLAlignmentKeyTests
{
    private static XLAlignmentKey MakeKey(
        XLAlignmentHorizontalValues horizontal = XLAlignmentHorizontalValues.Left,
        XLAlignmentVerticalValues vertical = XLAlignmentVerticalValues.Top,
        XLAlignmentReadingOrderValues readingOrder = XLAlignmentReadingOrderValues.LeftToRight)
    {
        return new XLAlignmentKey
        {
            Horizontal = horizontal,
            Vertical = vertical,
            Indent = 1,
            JustifyLastLine = false,
            ReadingOrder = readingOrder,
            RelativeIndent = 0,
            ShrinkToFit = false,
            TextRotation = 0,
            WrapText = false,
        };
    }

    [Test]
    public async Task Two_keys_equal_only_when_every_field_matches()
    {
        var a = MakeKey();
        var b = MakeKey();

        await Assert.That(a).IsEqualTo(b);
        await Assert.That(a.GetHashCode()).IsEqualTo(b.GetHashCode());

        await Assert.That(a).IsNotEqualTo(MakeKey(horizontal: XLAlignmentHorizontalValues.Right));
        await Assert.That(a).IsNotEqualTo(MakeKey(vertical: XLAlignmentVerticalValues.Bottom));
        await Assert.That(a).IsNotEqualTo(MakeKey(readingOrder: XLAlignmentReadingOrderValues.RightToLeft));
    }

    /// <summary>
    /// 999 does not land on any of the enum's 8 defined ordinals (0-7) under an unchecked byte
    /// narrowing (999 mod 256 = 231), so this is a value <see cref="Enum.IsDefined{TEnum}"/> can
    /// actually reject.
    /// </summary>
    [Test]
    public async Task Constructing_a_key_with_an_undefined_horizontal_value_throws()
    {
        var undefined = UncheckedCast(999);

        await Assert.That(() => MakeKey(horizontal: undefined)).Throws<ArgumentOutOfRangeException>();
    }

    /// <summary>
    /// 262 mod 256 is 6, which is <see cref="XLAlignmentHorizontalValues.Left"/> - a real, different,
    /// defined member. By the time any XLibur code sees this value it is, genuinely, <c>Left</c>:
    /// nothing distinguishes it from a caller who passed <c>Left</c> directly, because the wrap
    /// happens inside the caller's own cast expression, before the value is passed anywhere. This is
    /// unlike <c>XLBorderKey</c> and <c>XLColorKey</c>, where the narrowing that can wrap happens
    /// inside this library's own code, on a value still in its original, wider form - which is what
    /// lets those two catch it before the wrap occurs. This test exists to record the gap, not to
    /// call it acceptable.
    /// </summary>
    [Test]
    public async Task Wraparound_onto_a_defined_member_cannot_be_told_apart_from_that_member()
    {
        var wrapsToLeft = UncheckedCast(262);

        await Assert.That(wrapsToLeft).IsEqualTo(XLAlignmentHorizontalValues.Left);

        var key = MakeKey(horizontal: wrapsToLeft);

        await Assert.That(key.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Left);
    }

    [Test]
    public async Task Highest_defined_horizontal_value_does_not_throw()
    {
        // XLAlignmentHorizontalValues has 8 members, 0-7.
        var key = MakeKey(horizontal: XLAlignmentHorizontalValues.Right);

        await Assert.That(key.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Right);
    }

    [Test]
    public async Task Setting_an_undefined_horizontal_value_through_the_cell_facade_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var alignment = ws.Cell("A1").Style.Alignment;

        await Assert.That(() => alignment.Horizontal = UncheckedCast(999))
            .Throws<ArgumentOutOfRangeException>();
    }

    /// <summary>
    /// <c>XLDeferredAlignment</c> is a second, independent facade over the same key, reached through
    /// <c>IXLStyle.Batch</c> for a cell rather than through the ordinary <c>XLAlignment</c> facade.
    /// Validation lives on the key rather than on either facade precisely so a caller cannot reach
    /// an unguarded path through this one instead.
    /// </summary>
    [Test]
    public async Task Setting_an_undefined_horizontal_value_through_the_batch_facade_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell("A1");

        await Assert.That(() => cell.Style.Batch(s => s.Alignment.Horizontal = UncheckedCast(999)))
            .Throws<ArgumentOutOfRangeException>();
    }

    /// <summary>
    /// Stands in for a value computed at runtime rather than written as a literal. Unlike the
    /// int-backed <c>XLBorderStyleValues</c>, <c>XLColorType</c> and <c>XLThemeColor</c>, this enum's
    /// own byte-backed declaration makes an out-of-range literal cast a compile-time constant
    /// overflow the compiler already rejects (CS0221) - a second, independent guard, but one that
    /// cannot help against a value the compiler cannot see ahead of time, which is what a
    /// non-constant expression through this method reproduces.
    /// </summary>
    private static XLAlignmentHorizontalValues UncheckedCast(int value) => unchecked((XLAlignmentHorizontalValues)value);
}
