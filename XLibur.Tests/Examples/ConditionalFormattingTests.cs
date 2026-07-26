using XLibur.Examples.ConditionalFormatting;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class ConditionalFormattingTests
{
    [Test]
    public async Task CFColorScaleLowHigh()
    {
        await TestHelper.RunTestExample<CFColorScaleLowHigh>(@"ConditionalFormatting\CFColorScaleLowHigh.xlsx");
    }

    [Test]
    public async Task CFColorScaleLowMidHigh()
    {
        await TestHelper.RunTestExample<CFColorScaleLowMidHigh>(@"ConditionalFormatting\CFColorScaleLowMidHigh.xlsx");
    }

    [Test]
    public async Task CFColorScaleMinimumMaximum()
    {
        await TestHelper.RunTestExample<CFColorScaleMinimumMaximum>(@"ConditionalFormatting\CFColorScaleMinimumMaximum.xlsx");
    }

    [Test]
    public async Task CFContains()
    {
        await TestHelper.RunTestExample<CFContains>(@"ConditionalFormatting\CFContains.xlsx");
    }

    [Test]
    public async Task CFDataBar()
    {
        await TestHelper.RunTestExample<CFDataBar>(@"ConditionalFormatting\CFDataBar.xlsx");
    }

    [Test]
    public async Task CFDataBarNegative()
    {
        await TestHelper.RunTestExample<CFDataBarNegative>(@"ConditionalFormatting\CFDataBarNegative.xlsx");
    }

    [Test]
    public async Task CFEndsWith()
    {
        await TestHelper.RunTestExample<CFEndsWith>(@"ConditionalFormatting\CFEndsWith.xlsx");
    }

    [Test]
    public async Task CFEqualsNumber()
    {
        await TestHelper.RunTestExample<CFEqualsNumber>(@"ConditionalFormatting\CFEqualsNumber.xlsx");
    }

    [Test]
    public async Task CFEqualsString()
    {
        await TestHelper.RunTestExample<CFEqualsString>(@"ConditionalFormatting\CFEqualsString.xlsx");
    }

    [Test]
    public async Task CFIconSet()
    {
        await TestHelper.RunTestExample<CFIconSet>(@"ConditionalFormatting\CFIconSet.xlsx");
    }

    [Test]
    public async Task CFIsBlank()
    {
        await TestHelper.RunTestExample<CFIsBlank>(@"ConditionalFormatting\CFIsBlank.xlsx");
    }

    [Test]
    public async Task CFIsError()
    {
        await TestHelper.RunTestExample<CFIsError>(@"ConditionalFormatting\CFIsError.xlsx");
    }

    [Test]
    public async Task CFNotBlank()
    {
        await TestHelper.RunTestExample<CFNotBlank>(@"ConditionalFormatting\CFNotBlank.xlsx");
    }

    [Test]
    public async Task CFNotContains()
    {
        await TestHelper.RunTestExample<CFNotContains>(@"ConditionalFormatting\CFNotContains.xlsx");
    }

    [Test]
    public async Task CFNotEqualsNumber()
    {
        await TestHelper.RunTestExample<CFNotEqualsNumber>(@"ConditionalFormatting\CFNotEqualsNumber.xlsx");
    }

    [Test]
    public async Task CFNotEqualsString()
    {
        await TestHelper.RunTestExample<CFNotEqualsString>(@"ConditionalFormatting\CFNotEqualsString.xlsx");
    }

    [Test]
    public async Task CFNotError()
    {
        await TestHelper.RunTestExample<CFNotError>(@"ConditionalFormatting\CFNotError.xlsx");
    }

    [Test]
    public async Task CFStartsWith()
    {
        await TestHelper.RunTestExample<CFStartsWith>(@"ConditionalFormatting\CFStartsWith.xlsx");
    }

    [Test]
    public async Task CFMultipleConditions()
    {
        await TestHelper.RunTestExample<CFMultipleConditions>(@"ConditionalFormatting\CFMultipleConditions.xlsx");
    }

    [Test]
    public async Task CFStopIfTrue()
    {
        await TestHelper.RunTestExample<CFStopIfTrue>(@"ConditionalFormatting\CFStopIfTrue.xlsx");
    }

    [Test]
    public async Task CFTop()
    {
        await TestHelper.RunTestExample<CFTop>(@"ConditionalFormatting\CFTop.xlsx");
    }

    [Test]
    public async Task CFBottom()
    {
        await TestHelper.RunTestExample<CFBottom>(@"ConditionalFormatting\CFBottom.xlsx");
    }

    [Test]
    public async Task CFDatesOccurring()
    {
        await TestHelper.RunTestExample<CFDatesOccurring>(@"ConditionalFormatting\CFDatesOccurring.xlsx");
    }

    [Test]
    public async Task CFDataBars()
    {
        await TestHelper.RunTestExample<CFDataBars>(@"ConditionalFormatting\CFDataBars.xlsx");
    }
}
