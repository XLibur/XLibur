using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A formula typed through the API names the parameters of <c>LET</c> and <c>LAMBDA</c> as Excel
/// does in a file: <c>_xlpm.</c> before each parameter and each use of it, and <c>_xlop.</c> for an
/// optional parameter, written <c>[y]</c> when typed. Excel refuses to open a workbook whose
/// parameters lack the prefix.
/// </summary>
/// <remarks>
/// Every expected text is what Excel 365 wrote for the same formula typed into a cell over COM
/// (<c>Range.Formula2</c>), then saved.
/// </remarks>
public class LambdaParameterPrefixTests
{
    [Test]
    [Arguments("LET(x,1,x+1)", "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")]
    [Arguments("LAMBDA(x,x*2)(3)", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(3)")]
    [Arguments("LET(f,LAMBDA(x,x*2),f(3))", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(3))")]
    [Arguments("MAP(A1:A3,LAMBDA(v,v*2))", "_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))")]
    [Arguments("LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1)",
        "_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))(1)")]
    [Arguments("LET(x,1,LAMBDA(x,x+1)(x))", "_xlfn.LET(_xlpm.x,1,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(_xlpm.x))")]
    [Arguments("LET(x,1,x+Rate)", "_xlfn.LET(_xlpm.x,1,_xlpm.x+Rate)")]
    [Arguments("LET(r,A1,SUM(r:A3))", "_xlfn.LET(_xlpm.r,A1,SUM(_xlpm.r:A3))")]
    [Arguments("LET(a,1,b,a+1,a+b)", "_xlfn.LET(_xlpm.a,1,_xlpm.b,_xlpm.a+1,_xlpm.a+_xlpm.b)")]
    [Arguments("REDUCE(0,A1:A3,LAMBDA(acc,v,acc+v))", "_xlfn.REDUCE(0,A1:A3,_xlfn.LAMBDA(_xlpm.acc,_xlpm.v,_xlpm.acc+_xlpm.v))")]
    public async Task A_typed_formula_names_its_parameters_as_Excel_does(string typed, string excel)
    {
        await Assert.That(Typed(typed)).IsEqualTo(excel);
    }

    /// <summary>
    /// A call by the name of one of Excel's functions calls that function, even where a LET or
    /// LAMBDA binds the name: Excel 365 calculates <c>LET(SORT,LAMBDA(x,x*2),SORT(3))</c> as 3, and
    /// writes the call as <c>_xlfn._xlws.SORT(3)</c>. The name as a value is the parameter.
    /// </summary>
    [Test]
    [Arguments("LET(SORT,LAMBDA(x,x*2),SORT(3))", "_xlfn.LET(_xlpm.SORT,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlfn._xlws.SORT(3))")]
    [Arguments("LET(SUM,LAMBDA(x,x*2),SUM(3))", "_xlfn.LET(_xlpm.SUM,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),SUM(3))")]
    [Arguments("LET(CONCAT,5,CONCAT+1)", "_xlfn.LET(_xlpm.CONCAT,5,_xlpm.CONCAT+1)")]
    [Arguments("LET(sum,4,sum+1)", "_xlfn.LET(_xlpm.sum,4,_xlpm.sum+1)")]
    [Arguments("LAMBDA(SORT,SORT+1)(2)", "_xlfn.LAMBDA(_xlpm.SORT,_xlpm.SORT+1)(2)")]
    public async Task A_call_by_a_function_name_calls_the_function_and_not_the_parameter(string typed, string excel)
    {
        await Assert.That(Typed(typed)).IsEqualTo(excel);
    }

    [Test]
    [Arguments("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")]
    [Arguments("_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))(1)")]
    [Arguments("_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(3))")]
    public async Task A_formula_already_written_as_Excel_writes_it_is_left_as_it_is(string excel)
    {
        await Assert.That(Typed(excel)).IsEqualTo(excel);
    }

    [Test]
    [Arguments("LET(x,1,x)+x", "_xlfn.LET(_xlpm.x,1,_xlpm.x)+x")]
    [Arguments("LET(x,y,y,1,x+y)", "_xlfn.LET(_xlpm.x,y,_xlpm.y,1,_xlpm.x+_xlpm.y)")]
    [Arguments("LET(x,\"x\",x&\"x\")", "_xlfn.LET(_xlpm.x,\"x\",_xlpm.x&\"x\")")]
    [Arguments("LAMBDA(x,x)(x)", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(x)")]
    [Arguments("let(X,1,x+1)", "_xlfn.LET(_xlpm.X,1,_xlpm.x+1)")]
    public async Task Only_a_parameter_inside_its_scope_is_prefixed(string typed, string expected)
    {
        // A name before its LET binding, after the LET, or in the argument list of a call, is not
        // the parameter; nor is text in a string.
        await Assert.That(Typed(typed)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("LET(x,1,x+SUM(Sheet1!A1,A1:B2,Sheet1:Sheet2!A1,!A1,[1]Sheet1!A1,[1]Sheet1:Sheet2!A1))")]
    [Arguments("LET(x,1,IF(TRUE,#N/A,\"t\")&{1,2;3,4}&-x&(x)&x%&Sheet1!#REF!)")]
    [Arguments("LET(x,1,x+Rate+Sheet1!Rate+!Rate+[1]!Rate+[1]Sheet1!Rate)")]
    [Arguments("LET(x,1,x+Table1[Col]+SUM(Table1[[#Data],[Col]])+[1]!Table1[Col]+SUM(,x))")]
    [Arguments("LET(x,1,x+Sheet1!F(1)+[1]Sheet1!F(1)+[1]!F(1))")]
    [Arguments("LET(x,1,x&App|Topic!'item')")]
    [Arguments("LET(x,1,x&[1]!'item')")]
    public async Task Every_other_part_of_the_formula_is_left_as_it_is(string typed)
    {
        // Only the parameter x changes: references, names, tables, values, operators, functions of
        // a sheet or another workbook, and DDE keep their text.
        var written = Typed(typed);

        await Assert.That(written.Replace("_xlpm.x", "x").Replace("_xlfn.", "")).IsEqualTo(typed);
        await Assert.That(written).StartsWith("_xlfn.LET(_xlpm.x,1,");
    }

    [Test]
    [Arguments("LET(x,1)", "_xlfn.LET(x,1)")]
    [Arguments("LET(1,2,3)", "_xlfn.LET(1,2,3)")]
    [Arguments("LAMBDA(1,x)", "_xlfn.LAMBDA(1,x)")]
    [Arguments("LAMBDA()", "_xlfn.LAMBDA()")]
    [Arguments("CONCAT(x,\"LET\")", "_xlfn.CONCAT(x,\"LET\")")]
    [Arguments("CONCAT(A1)", "_xlfn.CONCAT(A1)")]
    public async Task A_call_that_declares_no_parameters_is_left_as_it_is(string typed, string expected)
    {
        // A LET or LAMBDA whose arguments are not a parameter list, as Excel would refuse, is read as
        // any other function; and a formula with neither has nothing to prefix.
        await Assert.That(Typed(typed)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("LET(x,1,LOG10(x))", "_xlfn.LET(_xlpm.x,1,LOG10(_xlpm.x))")]
    [Arguments("LET(x,1,B2(x))", "_xlfn.LET(_xlpm.x,1,B2(_xlpm.x))")]
    [Arguments("LET(x,{TRUE,\"t\",#N/A},x)", "_xlfn.LET(_xlpm.x,{TRUE,\"t\",#N/A},_xlpm.x)")]
    public async Task A_parameter_used_inside_any_part_is_prefixed(string typed, string expected)
    {
        // B2 reads as a cell called as a function; the array holds every kind of constant.
        await Assert.That(Typed(typed)).IsEqualTo(expected);
    }

    [Test]
    public async Task A_formula_typed_in_R1C1_names_its_parameters_too()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet().Cell("C1");
        cell.FormulaR1C1 = "LET(x,1,x+RC[-1])";

        await Assert.That(cell.FormulaA1).IsEqualTo("_xlfn.LET(_xlpm.x,1,_xlpm.x+B1)");
    }

    [Test]
    public async Task The_saved_file_holds_the_prefixed_formula()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet().Cell("C1").FormulaA1 = "LET(x,1,x+1)";
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(saved.Sheet1Xml()).Contains("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)");
    }

    private static string Typed(string formula)
    {
        // The workbook Excel typed the formulas in defines Rate, so Excel wrote it as that name, not as
        // its function RATE passed by name (_xleta.RATE).
        using var wb = new XLWorkbook();
        wb.DefinedNames.Add("Rate", "0.5");
        var cell = wb.AddWorksheet().Cell("C1");
        cell.FormulaA1 = formula;
        return cell.FormulaA1;
    }
}
