using XLibur.Examples.Styles;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class StylesTests
{
    [Test]
    public async Task DefaultStyles()
    {
        await TestHelper.RunTestExample<DefaultStyles>(@"Styles\DefaultStyles.xlsx");
    }

    [Test]
    public async Task PurpleWorksheet()
    {
        await TestHelper.RunTestExample<PurpleWorksheet>(@"Styles\PurpleWorksheet.xlsx");
    }

    [Test]
    public async Task StyleAlignment()
    {
        await TestHelper.RunTestExample<StyleAlignment>(@"Styles\StyleAlignment.xlsx");
    }

    [Test]
    public async Task StyleBorder()
    {
        await TestHelper.RunTestExample<StyleBorder>(@"Styles\StyleBorder.xlsx");
    }

    [Test]
    public async Task StyleFill()
    {
        await TestHelper.RunTestExample<StyleFill>(@"Styles\StyleFill.xlsx");
    }

    [Test]
    public async Task StyleFont()
    {
        await TestHelper.RunTestExample<StyleFont>(@"Styles\StyleFont.xlsx");
    }

    [Test]
    public async Task StyleNumberFormat()
    {
        await TestHelper.RunTestExample<StyleNumberFormat>(@"Styles\StyleNumberFormat.xlsx");
    }

    [Test]
    public async Task StyleIncludeQuotePrefix()
    {
        await TestHelper.RunTestExample<StyleIncludeQuotePrefix>(@"Styles\StyleIncludeQuotePrefix.xlsx");
    }

    [Test]
    public async Task StyleRowsColumns()
    {
        await TestHelper.RunTestExample<StyleRowsColumns>(@"Styles\StyleRowsColumns.xlsx");
    }

    [Test]
    public async Task StyleWorksheet()
    {
        await TestHelper.RunTestExample<StyleWorksheet>(@"Styles\StyleWorksheet.xlsx");
    }

    [Test]
    public async Task UsingColors()
    {
        await TestHelper.RunTestExample<UsingColors>(@"Styles\UsingColors.xlsx");
    }

    [Test]
    public async Task UsingPhonetics()
    {
        await TestHelper.RunTestExample<UsingPhonetics>(@"Styles\UsingPhonetics.xlsx");
    }

    [Test]
    public async Task UsingRichText()
    {
        await TestHelper.RunTestExample<UsingRichText>(@"Styles\UsingRichText.xlsx");
    }
}
