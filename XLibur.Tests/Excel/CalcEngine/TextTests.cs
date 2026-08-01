using XLibur.Excel;
using System;
using System.Globalization;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

[SetCulture("en-US")]
public class TextTests
{
    [Test]
    [Arguments("ABCDEF123", "ABCDEF123")]
    [Arguments("ァィゥェォッャュョヮ", "ｧｨｩｪｫｯｬｭｮヮ")] // Small katakana, there is no half wa variant
    [Arguments("アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン", "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ")]
    [Arguments("！＂＃\uff04％＆＇（）＊\uff0b，－．／０１２３４５６７８９：；\uff1c\uff1d\uff1e？＠", @"!""#$%&'()*+,-./0123456789:;<=>?@")]
    [Arguments("ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ", "ABCDEFGHIJKLMNOPQRSTUVWXYZ")]
    [Arguments("［＼］\uff3e＿\uff40ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛\uff5c｝\uff5e", @"[\]^_`abcdefghijklmnopqrstuvwxyz{|}~")]
    [Arguments(@"―‘’”、。「」゛゜・ー￥", @"ｰ`'""､｡｢｣ﾞﾟ･ｰ\")]
    public async Task Asc_converts_fullwidth_characters_to_halfwidth_characters(string input, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ASC(\"{input}\")")).IsEqualTo(expected);
    }

    [Test]
    public async Task Char_returns_error_on_empty_string()
    {
        // Calc engine tries to coerce it to number and fails. It never even reaches the functions.
        await Assert.That(XLWorkbook.EvaluateExpr(@"CHAR("""")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(0)]
    [Arguments(256)]
    [Arguments(9797)]
    public async Task Char_number_must_be_between_1_and_255(int number)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"CHAR({number})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(48, '0')]
    [Arguments(97, 'a')]
    [Arguments(128, '€')]
    [Arguments(138, 'Š')]
    [Arguments(169, '©')]
    [Arguments(182, '¶')]
    [Arguments(230, 'æ')]
    [Arguments(255, 'ÿ')]
    [Arguments(255.9, 'ÿ')]
    public async Task Char_interprets_number_as_win1252(double number, char expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"CHAR({number})");
        await Assert.That(actual).IsEqualTo(expected.ToString());
    }

    [Test]
    public async Task Clean_empty_string_is_empty_string()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"CLEAN("""")")).IsEqualTo("");
    }

    [Test]
    public async Task Clean_removes_control_characters()
    {
        var actual = XLWorkbook.EvaluateExpr(@"CLEAN(CHAR(9)&""Monthly report""&CHAR(10))");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("Monthly report"));

        actual = XLWorkbook.EvaluateExpr(@"CLEAN(""   "")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("   "));
    }

    [Test]
    public async Task Code_returns_error_on_empty_string()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"CODE("""")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("A", 65)]
    [Arguments("BCD", 66)]
    [Arguments("€", 128)]
    [Arguments("ÿ", 255)]
    public async Task Code_returns_win1252_codepoint_of_first_character(string text, int expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"CODE(\"{text}\")");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task Code_is_inverse_to_char()
    {
        for (var i = 1; i < 256; ++i)
            await Assert.That(XLWorkbook.EvaluateExpr($"CODE(CHAR({i}))")).IsEqualTo(i);
    }

    [Test]
    [Arguments("π")]
    [Arguments("ب")]
    [Arguments("😃")]
    [Arguments("♫")]
    [Arguments("ひ")]
    public async Task Code_returns_question_mark_code_on_non_win1252_chars(string text)
    {
        var expected = XLWorkbook.EvaluateExpr("CODE(\"?\")");
        var actual = XLWorkbook.EvaluateExpr($"CODE(\"{text}\")");
        await Assert.That(expected).IsEqualTo(63);
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Concat_concatenates_scalar_values()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var actual = ws.Evaluate(@"CONCAT(""ABC"",123,TRUE,IF(TRUE,),1.25)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABC123TRUE1,25"));

        actual = ws.Evaluate(@"CONCAT("""",""123"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("123"));

        ws.FirstCell().SetValue(20.5)
            .CellBelow().SetValue("AB")
            .CellBelow().SetFormulaA1("DATE(2019,1,1)")
            .CellBelow().SetFormulaA1("CONCAT(A1:A3)");

        actual = ws.Cell("A4").Value;
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("20,5AB43466"));
    }

    [Test]
    public async Task Concat_concatenates_array_values()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT({""A"",""B"",""C""},{0,1},{2;3},{4,5,6;7,8,9},""Z"")")).IsEqualTo("ABC0123456789Z");
    }

    [Test]
    public async Task Concat_concatenates_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C2").InsertData(new object[]
        {
            ("A", "B", "C"),
            (1, 2, 3, 4),
            (5, 6, 7, 8),
        });
        await Assert.That(ws.Evaluate("CONCAT(C2:E2,C3:F4,C2,\"Z\")")).IsEqualTo("ABC12345678AZ");
    }

    [Test]
    public async Task Concat_has_limit_of_32767_characters()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CONCAT(REPT(\"A\",32768))")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Concat_accepts_only_area_references()
    {
        // Only areas are accepted, not unions
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate("CONCAT((C2:E2,C3:F4),C2,\"Z\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Concat_propagates_error_values()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",#DIV/0!,5)")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",{""D"",#DIV/0!,7},5)")).IsEqualTo(XLError.DivisionByZero);

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B5").SetValue(XLError.DivisionByZero).CellBelow().SetValue(5);
        await Assert.That(ws.Evaluate("CONCAT(\"ABC\",B5:B6)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task Concat_treats_blanks_as_empty_string()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",,""123"",)")).IsEqualTo("ABC123");
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Concatenate_concatenates_scalar_values()
    {
        using var wb = new XLWorkbook();
        var actual = wb.Evaluate(@"CONCATENATE(""ABC"",123,4.56,IF(TRUE,),TRUE)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABC1234,56TRUE"));

        actual = wb.Evaluate(@"CONCATENATE("""",""123"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("123"));
    }

    [Test]
    public async Task Concatenate_with_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "Hello";
        ws.Cell("B1").Value = "World";
        ws.Cell("C1").FormulaA1 = "CONCATENATE(A1:A2,\" \",B1:B2)";
        ws.Cell("A3").FormulaA1 = "CONCATENATE(A1:A2,\" \",B1:B2)";

        await Assert.That(ws.Evaluate(@"CONCATENATE(A1,"" "",B1)")).IsEqualTo("Hello World");

        // The result on C1 is on the same row (only one intersected cell) means implicit intersection
        // results in a one value per intersection and thus correct value. The A3 intersects two cells
        // and thus results in #VALUE! error.
        await Assert.That(ws.Cell("C1").Value).IsEqualTo("Hello World");
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Concatenate_has_limit_of_32767_characters()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CONCATENATE(REPT(\"A\",32767))")).IsNotEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("CONCATENATE(REPT(\"A\",32768))")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Concatenate_uses_implicit_intersection_on_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue(20)
            .CellBelow().SetValue("AB")
            .CellBelow().SetFormulaA1("DATE(2019,1,1)");

        // Calling cell is 1st row, so formula should return A1
        ws.Cell("B1").SetFormulaA1("CONCATENATE(A1:A3)");
        await Assert.That(ws.Cell("B1").Value).IsEqualTo("20");

        // Calling cell is 2nd row, so formula should return A2
        ws.Cell("B2").SetFormulaA1("CONCATENATE(A1:A3)");
        await Assert.That(ws.Cell("B2").Value).IsEqualTo("AB");

        // Calling cell is 3rd row, so formula should return A3's textual representation
        ws.Cell("B3").SetFormulaA1("CONCATENATE(A1:A3)");
        await Assert.That(ws.Cell("B3").Value).IsEqualTo("43466");

        // Calling cell doesn't share row with any cell in parameter range.
        ws.Cell("A4").SetFormulaA1("CONCATENATE(A1:A3)");
        await Assert.That(ws.Cell("A4").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Dollar_coercion()
    {
        // Empty string is not coercible to number
        await Assert.That(XLWorkbook.EvaluateExpr("DOLLAR(\"\", 3)")).IsEqualTo(XLError.IncompatibleValue);
    }

    // en-US culture differs between .NET Fx and Core for negative currency -> no test for negative
    [Test]
    [Arguments(123.54, 3, "$123.540")]
    [Arguments(123.54, 3.9, "$123.540")]
    [Arguments(1234.567, 2, "$1,234.57")]
    [Arguments(1250, -2, "$1,300")]
    [Arguments(1, -1E+100, "$0")]
    public async Task Dollar_en(double number, double decimals, string expected)
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate($"DOLLAR({number},{decimals})").GetText()).IsEqualTo(expected);
    }

    [Test]
    [SetCulture("cs-CZ")]
    [Arguments(123.54, 3, "123,540 Kč")]
    [Arguments(-1234.567, 4, "-1 234,5670 Kč")]
    [Arguments(-1250, -2, "-1 300 Kč")]
    public async Task Dollar_cs(double number, double decimals, string expected)
    {
        using var wb = new XLWorkbook();
        var formula = $"DOLLAR({number.ToString(CultureInfo.InvariantCulture)},{decimals.ToString(CultureInfo.InvariantCulture)})";
        await Assert.That(wb.Evaluate(formula).GetText()).IsEqualTo(expected);
    }

    [Test]
    [SetCulture("de-DE")]
    [Arguments(1234.567, 2, "1.234,57 €")]
    [Arguments(1234.567, -2, "1.200 €")]
    [Arguments(-1234.567, 4, "-1.234,5670 €")]
    public async Task Dollar_de(double number, double decimals, string expected)
    {
        using var wb = new XLWorkbook();
        var formula = $"DOLLAR({number.ToString(CultureInfo.InvariantCulture)},{decimals.ToString(CultureInfo.InvariantCulture)})";
        await Assert.That(wb.Evaluate(formula).GetText()).IsEqualTo(expected);
    }

    [Test]
    public async Task Dollar_uses_two_decimal_places_by_default()
    {
        using var wb = new XLWorkbook();
        var actual = wb.Evaluate("DOLLAR(123.543)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("$123.54"));
    }

    [Test]
    public async Task Dollar_can_have_at_most_127_decimal_places()
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate("DOLLAR(1,99)")).IsEqualTo("$1." + new string('0', 99));
        await Assert.That(wb.Evaluate("DOLLAR(1,128)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Exact_Empty_Input_String()
    {
        Object actual = XLWorkbook.EvaluateExpr(@"Exact("""", """")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(true));
    }

    [Test]
    public async Task Exact_Value()
    {
        Object actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""asdf"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(true));

        actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""ASDF"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(false));

        actual = XLWorkbook.EvaluateExpr("Exact(123, 123)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(true));

        actual = XLWorkbook.EvaluateExpr("Exact(321, 123)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(false));
    }

    [Test]
    public async Task Find_Empty_Pattern_And_Empty_Text()
    {
        // Different behavior from SEARCH
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND("""", """")")).IsEqualTo(1);

        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND("""", ""a"", 2)")).IsEqualTo(2);
    }

    [Test]
    public async Task Find_Empty_Search_Pattern_Returns_Start_Of_Text()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND("""", ""asdf"")")).IsEqualTo(1);
    }

    [Test]
    public async Task Find_Looks_Only_From_Start_Position_Onward()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""This"", ""This is some text"", 2)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_Start_Position_Too_Large()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""abc"", ""abcdef"", 10)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_Start_Position_Too_Small()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""text"", ""This is some text"", 0)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_Empty_Searched_Text_Returns_Error()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""abc"", """")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_String_Not_Found()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""123"", ""asdf"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_Case_Sensitive_String_Not_Found()
    {
        // Find is case-sensitive
        await Assert.That(XLWorkbook.EvaluateExpr(@"FIND(""excel"", ""Microsoft Excel 2010"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_Value()
    {
        var actual = XLWorkbook.EvaluateExpr(@"FIND(""Tuesday"", ""Today is Tuesday"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(10));

        // Doesnt support wildcards
        actual = XLWorkbook.EvaluateExpr(@"FIND(""T*y"", ""Today is Tuesday"")");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Find_Arguments_Are_Converted_To_Expected_Types()
    {
        var actual = XLWorkbook.EvaluateExpr(@"FIND(1.2, ""A1.2B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(2));

        actual = XLWorkbook.EvaluateExpr(@"FIND(TRUE, ""ATRUE"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(2));

        actual = XLWorkbook.EvaluateExpr("FIND(23, 1.2345)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(3));

        actual = XLWorkbook.EvaluateExpr(@"FIND(""a"", ""aaaaa"", ""2 1/2"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(2));
    }

    [Test]
    public async Task Find_Error_Arguments_Return_The_Error()
    {
        var actual = XLWorkbook.EvaluateExpr(@"FIND(#N/A, ""a"")");
        await Assert.That(actual).IsEqualTo(XLError.NoValueAvailable);

        actual = XLWorkbook.EvaluateExpr(@"FIND("""", #N/A)");
        await Assert.That(actual).IsEqualTo(XLError.NoValueAvailable);

        actual = XLWorkbook.EvaluateExpr(@"FIND(""a"", ""a"", #N/A)");
        await Assert.That(actual).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task Fixed_coercion()
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate("""FIXED("asdf")""")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(wb.Evaluate("""FIXED(1234,1,"TRUE")""")).IsEqualTo("1234.0");
        await Assert.That(wb.Evaluate("""FIXED(1234,1,"FALSE")""")).IsEqualTo("1,234.0");
        await Assert.That(wb.Evaluate("""FIXED(1234,1,"0")""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Fixed_examples()
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate("FIXED(1234567)")).IsEqualTo("1,234,567.00");
        await Assert.That(wb.Evaluate("FIXED(1234567.555555,4,TRUE)")).IsEqualTo("1234567.5556");
        await Assert.That(wb.Evaluate("FIXED(.555555,10)")).IsEqualTo("0.5555550000");
        await Assert.That(wb.Evaluate("FIXED(1234567,-3)")).IsEqualTo("1,235,000");
    }

    [Test]
    public async Task Fixed_en()
    {
        var actual = XLWorkbook.EvaluateExpr("FIXED(17300.67,4)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("17,300.6700"));

        actual = XLWorkbook.EvaluateExpr("FIXED(17300.67,2,TRUE)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("17300.67"));

        actual = XLWorkbook.EvaluateExpr("FIXED(17300.67)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("17,300.67"));

        actual = XLWorkbook.EvaluateExpr("FIXED(1,-1E+300)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("0"));
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Fixed_cs()
    {
        using var wb = new XLWorkbook();
        var actual = wb.Evaluate("FIXED(17300.67,4)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("17 300,6700"));

        actual = wb.Evaluate("FIXED(17300.67,2,TRUE)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("17300,67"));

        actual = wb.Evaluate("FIXED(17300.67)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("17 300,67"));
    }

    [Test]
    public async Task Fixed_can_have_at_most_127_decimal_places()
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate("FIXED(1,99)")).IsEqualTo("1." + new string('0', 99));
        await Assert.That(wb.Evaluate("FIXED(1,128)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Left_returns_whole_text_when_requested_length_is_greater_than_text_length()
    {
        var actual = XLWorkbook.EvaluateExpr(@"LEFT(""ABC"", 5)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABC"));
    }

    [Test]
    public async Task Left_takes_one_character_by_default()
    {
        var actual = XLWorkbook.EvaluateExpr("""LEFT("ABC")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("A"));
    }

    [Test]
    public async Task Left_returns_error_on_negative_number_of_chars()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("""LEFT("ABC", -1)""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Left_returns_empty_string_on_empty_input()
    {
        var actual = XLWorkbook.EvaluateExpr("""LEFT("")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    [Arguments("ABC", 2, "AB")]
    [Arguments("ABC", 2.9, "AB")]
    [Arguments("ABC", 3, "ABC")]
    [Arguments("\uD83D\uDC69Z", 1, "\uD83D\uDC69")] // Paired surrogate
    [Arguments("\uD83D\uDC69Z", 2, "\uD83D\uDC69Z")] // Paired surrogate
    public async Task Left_takes_specified_number_of_characters(string text, double numChars, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""LEFT("{text}", {numChars})""").GetText()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("", 0)]
    [Arguments("word", 4)]
    [Arguments("A\r\n", 3)]
    [Arguments("H", 1)]
    [Arguments("\ud83d\ude0a", 2)] // Smile emoji
    [Arguments("Smile: \ud83d\ude0a!", 10)] // Smile emoji
    public async Task Len_returns_number_of_code_units(string text, double expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""LEN("{text}")""").GetNumber()).IsEqualTo(expected);
    }

    [Test]
    [SetCulture("en-US")]
    [Arguments("", "")]
    [Arguments("ABC", "abc")]
    [Arguments("Intelligence 2.0!", "intelligence 2.0!")]
    [Arguments("ͶꝎＫǢ", "ͷꝏｋǣ")] // Converts even non-latin chars
    [Arguments("Σ SUM Σ end Σ", "σ sum σ end ς")] // Bug for bug behavior of Excel. Σ at the end is turned to ς
    public async Task Lower_en(string text, string expected)
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate($"""LOWER("{text}")""").GetText()).IsEqualTo(expected);
    }

    [Test]
    [SetCulture("tr-TR")]
    [Arguments("INTELLIGENCE 2.0!", "ıntellıgence 2.0!")] // Turkey converts I to i without dot
    [Arguments("ΣΣΣΣ", "σσσς")]
    public async Task Lower_tr(string text, string expected)
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate($"""LOWER("{text}")""").GetText()).IsEqualTo(expected);
    }

    [Test]
    public async Task Mid_returns_rest_of_text_when_end_is_out_of_text_bounds()
    {
        var actual = XLWorkbook.EvaluateExpr("""MID("ABC",1,5)""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABC"));
    }

    [Test]
    public async Task Mid_when_start_is_after_end_of_text_return_empty_string()
    {
        var actual = XLWorkbook.EvaluateExpr("""MID("ABC",5,5)""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    [Arguments(0.9)]
    [Arguments(0)]
    [Arguments(-5)]
    [Arguments(int.MaxValue + 1d)]
    [Arguments(int.MaxValue + 5d)]
    public async Task Mid_start_must_be_at_least_one_and_at_most_max_int(double start)
    {
        var actual = XLWorkbook.EvaluateExpr($"""MID("ABC",{start},1)""");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(-0.1)]
    [Arguments(-5)]
    [Arguments(int.MaxValue + 1d)]
    [Arguments(int.MaxValue + 5d)]
    public async Task Mid_length_must_be_at_least_zero_and_at_most_max_int(double length)
    {
        var actual = XLWorkbook.EvaluateExpr($"""MID("ABC",1,{length})""");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("", 1, 1, "")]
    [Arguments("ABC", 2, 2, "BC")]
    [Arguments("ABC", 2, 0, "")]
    [Arguments("ABC", 3, 5, "")]
    [Arguments("abcdef", 3, 2, "cd")]
    [Arguments("abcdef", 4, 5, "def")]
    public async Task Mid_returns_substring(string text, double start, double length, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""MID("{text}",{start},{length})""").GetText()).IsEqualTo(expected);
    }

    [Test]
    public async Task Mid_uses_code_units()
    {
        // MID returns unpaired surrogates
        await Assert.That(XLWorkbook.EvaluateExpr("""MID("😊😊😊",1,3)""")).IsEqualTo("😊\uD83D");
        await Assert.That(XLWorkbook.EvaluateExpr("""MID("😊😊😊",1,4)""")).IsEqualTo("😊😊");
        await Assert.That(XLWorkbook.EvaluateExpr("""MID("😊😊😊",2,4)""")).IsEqualTo("\uDE0A😊\uD83D");
        await Assert.That(XLWorkbook.EvaluateExpr("""LEN(MID("😊😊😊",1,3))""")).IsEqualTo(3);
    }

    [Test]
    [Arguments("", 0d)]
    [Arguments("+ 1", 1d)]
    [Arguments("+1", 1d)]
    [Arguments("+1.23", 1.23)]
    [Arguments("- 1.23", -1.23)]
    [Arguments(" - 0 1 2 . 3 4 ", -12.34)]
    [Arguments(" - 0 \t1\t2\r .\n3 4 ", -12.34)]
    [Arguments(".1", 0.1)]
    [Arguments("-.1", -0.1)]
    [Arguments("1.234567890E+307", 1.234567890E+307)]
    [Arguments("1.234567890E-307", 1.234567890E-307d)]
    [Arguments("1.234567890E-309", 0d)]
    [Arguments("-1.234567890E-307", -1.234567890E-307d)]
    [Arguments(".99999999999999", 0.99999999999999)]
    [Arguments("1,23,4", 1234)]
    [Arguments("1,234,56", 123456)]
    [Arguments("1e-308", 0)]
    [Arguments("-1e-308", 0)]
    [Arguments("75825%", 758.25)]
    [Arguments("75825%%", 7.5825)]
    [Arguments("(56.4)", -56.4)]
    [Arguments("(128)%", -1.28)]
    public async Task NumberValue_converts_text_to_number(string text, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExprCurrent($"NUMBERVALUE(\"{text}\")");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    [SetCulture("de-DE")]
    public async Task NumberValue_takes_separators_from_current_culture()
    {
        var actual = (double)XLWorkbook.EvaluateExprCurrent("NUMBERVALUE(\"10.0.00.0,25\")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(100000.25));
    }

    [Test]
    [Arguments("1,234.56", ".", ",", 1234.56d)]
    [Arguments("1.234,56", ",", ".", 1234.56d)]
    [Arguments("1.234,56", ",ABC", ".DEF", 1234.56d)] // Only first char of separators is used
    public async Task NumberValue_optional_parameters_can_set_decimal_and_group_separators(string text, string @decimal, string group, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"NUMBERVALUE(\"{text}\",\"{@decimal}\",\"{group}\")");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments("NUMBERVALUE(\"123.45\", \".\", \".\")")] // Group separator same as decimal separator
    [Arguments("NUMBERVALUE(\"1.234.5\")")] // Two decimal separators
    [Arguments("NUMBERVALUE(\"1.234,5\")")] // Decimal separator before group separator
    [Arguments("NUMBERVALUE(\"12;34\")")] // Illegal character
    [Arguments("NUMBERVALUE(\"--1\")")] // Two minuses
    [Arguments("NUMBERVALUE(\"1.234567890E+308\")")] // Too large
    [Arguments("NUMBERVALUE(\"-1.234567890E+308\")")] // Too large (negative)
    [Arguments("NUMBERVALUE(\"1.234567890E-310\")")] // Too tiny
    [Arguments("NUMBERVALUE(\"-1.234567890E-310\")")] // Too tiny (negative)
    [Arguments("NUMBERVALUE(\"1\",\".\",\"\")")] // Empty group separator
    [Arguments("NUMBERVALUE(\"1\",\"\",\",\")")] // Empty decimal separators
    public async Task NumberValue_returns_error_on_unparsable_texts_out_of_range(string expression)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(expression)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("", "")]
    [Arguments("12aBC d123aD#$%sd^", "12Abc D123Ad#$%Sd^")]
    [Arguments("this is a TITLE", "This Is A Title")]
    [Arguments("2-way street", "2-Way Street")]
    [Arguments("76BudGet", "76Budget")]
    [Arguments("my name is francois botha", "My Name Is Francois Botha")]
    [Arguments("\ud83a\udd32", "\ud83a\udd32")] // U+1E932 has uppercase variant, but nothing changes, because PROPER uses code units
    public async Task Proper_upper_cases_first_letter_and_lower_cases_next_letters(string text, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""PROPER("{text}")""").GetText()).IsEqualTo(expected);
    }

    [Test]
    [Arguments(1, 1)]
    [Arguments(1, 0)]
    [Arguments(1, 10)]
    [Arguments(10, 1)]
    [Arguments(10, 10)]
    public async Task Replace_beyond_limit_appends_replacement(int startPos, int length)
    {
        var actual = XLWorkbook.EvaluateExpr($"""REPLACE("",{startPos},{length},"new text")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("new text"));
    }

    [Test]
    [Arguments("Here is some obsolete text to replace.", 14, 13, "new text", "Here is some new text to replace.")]
    [Arguments("ABC", 1, 2, "D", "DC")]
    [Arguments("ABC", 3, 1, "D", "ABD")]
    [Arguments("ABC", 3, 0, "D", "ABDC")]
    [Arguments("ABC", 4, 1, "D", "ABCD")]
    [Arguments("ABC", 4, 0, "D", "ABCD")]
    [Arguments("ABC", 1, 3, "D", "D")]
    [Arguments("ABC", 2, 2, "D", "AD")]
    [Arguments("ABC", 2, 0, "D", "ADBC")]
    [Arguments("ABC", 2, 3, "D", "AD")]
    [Arguments("abcdefghijk", 3, 4, "XY", "abXYghijk")]
    [Arguments("abcdefghijk", 3, 1, "12345", "ab12345defghijk")]
    [Arguments("abcdefghijk", 15, 4, "XY", "abcdefghijkXY")]
    public async Task Replace_replaces_value(string text, double startPos, int length, string replacement, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""REPLACE("{text}",{startPos},{length},"{replacement}")""").GetText()).IsEqualTo(expected);
    }

    [Test]
    public async Task Replace_start_position_must_be_from_1_to_32767()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,0,"D")""")).IsEqualTo("DABC");
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",0.9,0,"D")""")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",-1,0,"D")""")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,32767.9,"D")""")).IsEqualTo("D");
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,32768,"D")""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Replace_length_must_be_from_0_to_32767()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,0,"")""")).IsEqualTo("ABC");
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,-0.1,"D")""")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1, 32767.9,"D")""")).IsEqualTo("D");
        await Assert.That(XLWorkbook.EvaluateExpr("""REPLACE("ABC",1, 32768,"D")""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Rept_returns_empty_string_when_text_is_empty_string()
    {
        var actual = XLWorkbook.EvaluateExpr("""REPT("",3)""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    [Arguments(-1)]
    [Arguments(-0.1)]
    [Arguments(2147483648)]
    public async Task Rept_returns_error_when_count_is_negative_or_greater_than_max_int(double count)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""REPT("",{count})""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Rept_limits_output_text_length_to_32767()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("""REPT("A",32767)""")).IsEqualTo(new string('A', 32767));
        await Assert.That(XLWorkbook.EvaluateExpr("""REPT("A",32768)""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("ABC", 3, "ABCABCABC")]
    [Arguments("123", 2.5, "123123")]
    [Arguments("Francois", 0, "")]
    [Arguments("Francois Botha,", 3, "Francois Botha,Francois Botha,Francois Botha,")]
    public async Task Rept_Value(string text, double count, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""REPT("{text}",{count})""").GetText()).IsEqualTo(expected);
    }

    [Test]
    [Arguments(5)]
    [Arguments(3)]
    public async Task Right_returns_whole_text_when_requested_length_is_greater_than_text_length(int length)
    {
        var actual = XLWorkbook.EvaluateExpr($"""RIGHT("ABC",{length})""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABC"));
    }

    [Test]
    public async Task Right_takes_one_character_by_default()
    {
        var actual = XLWorkbook.EvaluateExpr("""RIGHT("ABC")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("C"));
    }

    [Test]
    public async Task Right_returns_error_on_negative_number_of_chars()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("""RIGHT("ABC",-1)""")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Right_returns_empty_string_on_empty_input()
    {
        var actual = XLWorkbook.EvaluateExpr("""RIGHT("")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    [Arguments("ABC", 0, "")]
    [Arguments("ABC", 1, "C")]
    [Arguments("ABC", 2, "BC")]
    [Arguments("ABC", 3, "ABC")]
    [Arguments("ABC", 4, "ABC")]
    [Arguments("ABC", 2.9, "BC")]
    [Arguments("Z\uD83D\uDC69", 1, "\uD83D\uDC69")] // Smiley emoji
    [Arguments("\uD83D\uDC69Z", 2, "\uD83D\uDC69Z")]
    [Arguments("\uD83D\uDC69Z", 3, "\uD83D\uDC69Z")]
    public async Task Right_takes_specified_number_of_characters(string text, double numChars, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""RIGHT("{text}",{numChars})""").GetText()).IsEqualTo(expected);
    }

    [Test]
    public async Task Search_Empty_Pattern_And_Empty_Text()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH("""", """")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Empty_Search_Pattern_Returns_Start_Of_Text()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SEARCH("""", ""asdf"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(1));
    }

    [Test]
    public async Task Search_Looks_Only_From_Start_Position_Onward()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""This"", ""This is some text"", 2)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Start_Position_Too_Large()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""abc"", ""abcdef"", 10)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Start_Position_Too_Small()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""text"", ""This is some text"", 0)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Empty_Searched_Text_Returns_Error()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""abc"", """")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Text_Not_Found()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""123"", ""asdf"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Wildcard_String_Not_Found()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"SEARCH(""soft?2010"", ""Microsoft Excel 2010"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    // http://www.excel-easy.com/examples/find-vs-search.html
    [Test]
    public async Task Search_Value()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SEARCH(""Tuesday"", ""Today is Tuesday"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(10));

        // The search is case-insensitive
        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""excel"", ""Microsoft Excel 2010"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(11));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""soft*2010"", ""Microsoft Excel 2010"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(6));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""Excel 20??"", ""Microsoft Excel 2010"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(11));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""text"", ""This is some text"", 14)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(14));
    }

    [Test]
    public async Task Search_Tilde_Escapes_Next_Char()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SEARCH(""~a~b~"", ""ab"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(1));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~*"", ""a*"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(1));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~*"", ""ab"")");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~?"", ""a?"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(1));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~?"", ""ab"")");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Search_Arguments_Are_Converted_To_Expected_Types()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SEARCH(1.2, ""A1.2B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(2));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(TRUE, ""ATRUE"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(2));

        actual = XLWorkbook.EvaluateExpr("SEARCH(23, 1.2345)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(3));

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a"", ""aaaaa"", ""2 1/2"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(2));
    }

    [Test]
    public async Task Search_Error_Arguments_Return_The_Error()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SEARCH(#N/A, ""a"")");
        await Assert.That(actual).IsEqualTo(XLError.NoValueAvailable);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH("""", #N/A)");
        await Assert.That(actual).IsEqualTo(XLError.NoValueAvailable);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a"", ""a"", #N/A)");
        await Assert.That(actual).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task Substitute_replaces_n_th_occurence()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday."", ""Tuesday"", ""Monday"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("This is a Monday."));

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 1)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("This is a Monday. Next week also has a Tuesday."));

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 2)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("This is a Tuesday. Next week also has a Monday."));

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", """", ""Monday"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("This is a Tuesday. Next week also has a Tuesday."));

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", """")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("This is a . Next week also has a ."));
    }

    [Test]
    public async Task Substitute_on_empty_string_returns_empty_string()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE("""","""",""Monday"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    public async Task Substitute_is_case_sensitive()
    {
        var actual = XLWorkbook.EvaluateExpr("""SUBSTITUTE("A","a","Z")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("A"));
    }

    [Test]
    public async Task Substitute_returns_original_string_when_occurrence_is_not_found()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABCABC"",""A"",""Z"",3)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABCABC"));
    }

    [Test]
    public async Task Substitute_searches_for_every_occurence()
    {
        // AA is matches at every character, it doesn't skip
        var actual = XLWorkbook.EvaluateExpr("""SUBSTITUTE("AAAAAAAA","AA","ZZ",3)""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("AAZZAAAA"));
    }

    [Test]
    public async Task Substitute_occurence_must_be_between_one_and_max_int()
    {
        var actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"",0.9)");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"", 2147483646.9)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABC"));

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"", 2147483647)");
        await Assert.That(actual).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task T_returns_empty_string_on_non_text()
    {
        var actual = XLWorkbook.EvaluateExpr("T(TODAY())");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));

        actual = XLWorkbook.EvaluateExpr("T(IF(TRUE,,))");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));

        actual = XLWorkbook.EvaluateExpr("T(TRUE)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));

        actual = XLWorkbook.EvaluateExpr("T(123)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    public async Task T_propagates_error()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("T(#DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task T_returns_text_when_value_is_text()
    {
        var actual = XLWorkbook.EvaluateExpr("""T("asdf")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("asdf"));

        actual = XLWorkbook.EvaluateExpr("""T("")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(""));
    }

    [Test]
    public async Task T_returns_array_of_results_when_argument_is_array()
    {
        const string formula = """T({"A",5,"B"})""";
        await Assert.That(XLWorkbook.EvaluateExpr($"""COLUMNS({formula})""")).IsEqualTo(3);
        await Assert.That(XLWorkbook.EvaluateExpr($"""ROWS({formula})""")).IsEqualTo(1);
        await Assert.That(XLWorkbook.EvaluateExpr($"""INDEX({formula},1,1)""")).IsEqualTo("A");
        await Assert.That(XLWorkbook.EvaluateExpr($"""INDEX({formula},1,2)""")).IsEqualTo("");
        await Assert.That(XLWorkbook.EvaluateExpr($"""INDEX({formula},1,3)""")).IsEqualTo("B");

        // Array doesn't propagate single error, but returns errors in the array
        await Assert.That(XLWorkbook.EvaluateExpr("""INDEX(T({"A",#REF!}),1,1)""")).IsEqualTo("A");
        await Assert.That(XLWorkbook.EvaluateExpr("""INDEX(T({"A",#REF!}),1,2)""")).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task T_returns_text_of_first_cell_in_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = "ABC";
        ws.Cell("B4").Value = 10;
        ws.Cell("B5").Value = XLError.NoValueAvailable;

        await Assert.That(ws.Evaluate("T(B3:B4)")).IsEqualTo("ABC");
        await Assert.That(ws.Evaluate("TYPE(T(B3:B4))")).IsEqualTo(2); // Is text, not array

        await Assert.That(ws.Evaluate("T(B4:C4)")).IsEqualTo(string.Empty);

        await Assert.That(ws.Evaluate("T(B5:C5)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task Text_returns_empty_string_on_empty_string()
    {
        var actual = XLWorkbook.EvaluateExpr(@"TEXT(1913415.93,"""")");
        await Assert.That(actual).IsEqualTo(string.Empty);
    }

    [Test]
    [Arguments("DATE(2010, 1, 1)", "yyyy-MM-dd", "2010-01-01")]
    [Arguments("1469.07", "0,000,000.00", "0,001,469.07")]
    [Arguments("1913415.93", "#,000.00", "1,913,415.93")]
    [Arguments("2800", "$0.00", "$2800.00")]
    [Arguments("0.4", "0%", "40%")]
    [Arguments("DATE(2010, 1, 1)", "MMMM yyyy", "January 2010")]
    [Arguments("DATE(2010, 1, 1)", "M/d/y", "1/1/10")]
    [Arguments("1234.567", "$0.00", "$1234.57")]
    [Arguments(".125", "$0.0%", "$12.5%")]
    [Arguments("1234.567", "YYYY-MM-DD HH:MM:SS", "1903-05-18 13:36:28")] // Excel is one second off (29), but that is in the library
    [Arguments("\"0.0245\"", "00%", "02%")]
    public async Task Text_formats_number(string numberArg, string format, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"TEXT({numberArg},\"{format}\")").GetText()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"211x\"", "211x")]
    [Arguments("true", "TRUE")]
    public async Task Text_returns_string_representation_of_non_numbers(string valueArg, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($@"TEXT({valueArg},""#00"")").GetText()).IsEqualTo(expected);
    }

    [Test]
    [Arguments(2020, 11, 1, 9, 23, 11, "m/d/yyyy h:mm:ss", "11/1/2020 9:23:11")]
    [Arguments(2023, 7, 14, 2, 12, 3, "m/d/yyyy h:mm:ss", "7/14/2023 2:12:03")]
    [Arguments(2025, 10, 14, 2, 48, 55, "m/d/yyyy h:mm:ss", "10/14/2025 2:48:55")]
    [Arguments(2023, 2, 19, 22, 1, 38, "m/d/yyyy h:mm:ss", "2/19/2023 22:01:38")]
    [Arguments(2025, 12, 19, 19, 43, 58, "m/d/yyyy h:mm:ss", "12/19/2025 19:43:58")]
    [Arguments(2034, 11, 16, 1, 48, 9, "m/d/yyyy h:mm:ss", "11/16/2034 1:48:09")]
    [Arguments(2018, 12, 10, 11, 22, 42, "m/d/yyyy h:mm:ss", "12/10/2018 11:22:42")]
    public async Task Text_formats_serial_dates(int year, int months, int days, int hour, int minutes, int seconds, string format, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($@"TEXT(DATE({year},{months},{days}) + TIME({hour},{minutes},{seconds}),""{format}"")")).IsEqualTo(expected);
    }

    [Test]
    public async Task Text_propagates_errors()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"TEXT(#REF!,""#00"")")).IsEqualTo(XLError.CellReference);
    }

    [Test]
    [Arguments("TEXTJOIN(\",\",TRUE,A1:B2)", "A,B,D")]
    [Arguments("TEXTJOIN(\",\",FALSE,A1:B2)", "A,,B,D")]
    [Arguments("TEXTJOIN(\",\",FALSE,A1,A2,B1,B2)", "A,B,,D")]
    [Arguments("TEXTJOIN(\",\",FALSE,1)", "1")]
    [Arguments("TEXTJOIN(\",\", TRUE, A:A, B:B)", "A,B,D")]
    [Arguments("TEXTJOIN(\",\", TRUE, D1:E2)", "")]
    [Arguments("TEXTJOIN(\",\", FALSE, D1:E2)", ",,,")]
    [Arguments("TEXTJOIN(\",\", FALSE, D1:D32768)", ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,")]
    [Arguments("TEXTJOIN(0, FALSE, A1:B2)", "A00B0D")]
    [Arguments("TEXTJOIN(false, FALSE, A1:B2)", "AFALSEFALSEBFALSED")]
    [Arguments("TEXTJOIN(\",\", 0, A1:B2)", "A,,B,D")]
    [Arguments("TEXTJOIN(\",\", 100, A1:B2)", "A,B,D")]
    [Arguments("TEXTJOIN(B2, FALSE, A1:B2)", "ADDBDD")]
    [Arguments("TEXTJOIN(\",\", FALSE, 12345.67, DATE(2018, 10, 30))", "12345.67,43403")]
    [Arguments("TEXTJOIN(\",\", \"FALSE\", A1:B2)", "A,,B,D")]
    public async Task TextJoin_joins_arguments_with_specified_delimiter(string formula, string expectedOutput)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "A";
        ws.Cell("A2").Value = "B";
        ws.Cell("B1").Value = "";
        ws.Cell("B2").Value = "D";

        ws.Cell("C1").FormulaA1 = formula;
        var a = ws.Cell("C1").Value;

        await Assert.That(a).IsEqualTo(expectedOutput);
    }

    [Test]
    [Arguments("TEXTJOIN(\",\", FALSE, D1:D32769)")]
    public async Task TextJoin_output_can_be_at_most_32767(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("C1").FormulaA1 = formula;

        // Excel actually returns #CALC!, but we don't have that error, mostly
        // because parser doesn't recognize it.
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("TEXTJOIN(\",\", \"Invalid\", \"Hello\", \"World\")")]
    public async Task TextJoin_coercion(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("", "")]
    [Arguments(" ", "")]
    [Arguments("    ", "")]
    [Arguments(" Break\r\n   Line   ", "Break\r\n Line")]
    [Arguments("non-whitespace-text", "non-whitespace-text")]
    [Arguments("white space text", "white space text")]
    [Arguments(" some text with padding   ", "some text with padding")]
    [Arguments(" \t  A  \t ", "\t A \t")]
    public async Task Trim_trims_spaces_and_removes_multi_spaces_from_inside_text(string text, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"""TRIM("{text}")""").GetText()).IsEqualTo(expected);
    }

    [Test]
    public async Task Upper_empty_string_returns_empty_string()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("""UPPER("")""")).IsEqualTo("");
    }

    [Test]
    public async Task Upper_converts_text_to_upper_case()
    {
        var actual = XLWorkbook.EvaluateExpr("""UPPER("AbCdEfG")""");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("ABCDEFG"));
    }

    [SetCulture("tr-TR")]
    [Test]
    public async Task Upper_uses_workbook_culture()
    {
        // Türkiye converts i to İ, not I.
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate("""UPPER("intelligence 2.0!")""")).IsEqualTo("İNTELLİGENCE 2.0!");
    }

    [Test]
    public async Task Value_Input_String_Is_Not_A_Number()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"VALUE(""asdf"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Value_FromBlankIsZero()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate("VALUE(A1)")).IsEqualTo(0d);
    }

    [Test]
    public async Task Value_FromEmptyStringIsError()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("VALUE(\"\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Value_PassingUnexpectedTypes()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("VALUE(14)")).IsEqualTo(14d);
        await Assert.That(XLWorkbook.EvaluateExpr("VALUE(TRUE)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("VALUE(FALSE)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("VALUE(#DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task Value_Value()
    {
        using var wb = new XLWorkbook();

        // Examples from spec
        await Assert.That(wb.Evaluate("VALUE(\"123.456\")")).IsEqualTo(123.456d);
        await Assert.That(wb.Evaluate("VALUE(\"$1,000\")")).IsEqualTo(1000d);
        await Assert.That(wb.Evaluate("VALUE(\"23-Mar-2002\")")).IsEqualTo(new DateTime(2002, 3, 23, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());
        await Assert.That((double)wb.Evaluate("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")")).IsEqualTo(0.188056d).Within(0.000001d);
    }

    [Test]
    [SetCulture("cs-CZ")]
    public async Task Value_NonEnglish()
    {
        using var wb = new XLWorkbook();

        // Examples from spec
        await Assert.That(wb.Evaluate("VALUE(\"123,456\")")).IsEqualTo(123.456d);
        await Assert.That(wb.Evaluate("VALUE(\"1 000 Kč\")")).IsEqualTo(1000d);
        await Assert.That(wb.Evaluate("VALUE(\"23-bře-2002\")")).IsEqualTo(37338d);
        await Assert.That((double)wb.Evaluate("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")")).IsEqualTo(0.188056d).Within(0.000001d);

        // Various number/currency formats
        await Assert.That(wb.Evaluate("VALUE(\"(1)\")")).IsEqualTo(-1d);
        await Assert.That(wb.Evaluate("VALUE(\"(100%)\")")).IsEqualTo(-1d);
        await Assert.That(wb.Evaluate("VALUE(\"(100%)\")")).IsEqualTo(-1d);
        await Assert.That(wb.Evaluate("VALUE(\"(1,5e1 Kč)\")")).IsEqualTo(-15d);
        await Assert.That(wb.Evaluate("VALUE(\"(1,5e3%)\")")).IsEqualTo(-15d);
        await Assert.That(wb.Evaluate("VALUE(\"(1,5e3)%\")")).IsEqualTo(-15d);

        var expectedSerialDate = new DateTime(2022, 3, 5, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime();
        await Assert.That(wb.Evaluate("VALUE(\"5-březen-22\")")).IsEqualTo(expectedSerialDate);
        await Assert.That(wb.Evaluate("VALUE(\"05.03.2022\")")).IsEqualTo(expectedSerialDate);
        await Assert.That(wb.Evaluate("VALUE(\"5-březen\")")).IsEqualTo(new DateTime(DateTime.Now.Year, 3, 5, 0, 0, 0, DateTimeKind.Unspecified).ToSerialDateTime());
    }
}
