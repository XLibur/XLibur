using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// One shared formula from a file: the R1C1 text that every cell of it was loaded from, and the tree
/// parsed from that text once for all of them.
/// </summary>
/// <remarks>
/// <para>
/// Excel saves a formula filled down a column as one shared formula. The loader gives each of its
/// cells A1 text of its own, and this object, which the cell keeps until its formula changes or moves
/// (<see cref="XLCellFormula.TryGetShared"/>). Evaluation and the dependency tree read the cell through
/// the one tree here, instead of parsing the A1 text of each cell (#513, #686). A relative reference
/// in the tree resolves against the cell it is evaluated for.
/// </para>
/// <para>
/// The tree lives as long as some cell still keeps this object. It belongs to the engine that parsed
/// it, like the tree a cell keeps of its own text (<see cref="XLCellFormula.GetAst"/>).
/// </para>
/// </remarks>
internal sealed class SharedFormulaGroup(string r1c1)
{
    private Formula? _ast;
    private bool _parsed;

    /// <summary>The R1C1 text of the formula.</summary>
    internal string R1C1 { get; } = r1c1;

    /// <summary>
    /// Get the tree of <see cref="R1C1"/>. The first call parses it, and later calls give the same tree.
    /// </summary>
    /// <remarks>
    /// When the parser refuses the text, each cell reads its own A1 text instead, so that a refusal is
    /// always the one the A1 text gets. The parser also refuses a cell called as a function, which reads
    /// a different cell in each cell of the formula.
    /// </remarks>
    /// <param name="engine">The engine of the workbook the formula is in.</param>
    /// <param name="ast">The tree, when the parser accepted the text.</param>
    /// <returns><c>false</c> when the parser refused the text.</returns>
    internal bool TryGetAst(XLCalcEngine engine, [NotNullWhen(true)] out Formula? ast)
    {
        if (!_parsed)
        {
            engine.TryParseR1C1(R1C1, out _ast);
            _parsed = true;
        }

        ast = _ast;
        return ast is not null;
    }
}
