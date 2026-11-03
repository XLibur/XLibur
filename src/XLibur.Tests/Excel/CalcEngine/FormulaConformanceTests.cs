using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static XLibur.Tests.Excel.CalcEngine.FormulaConformance;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Spec 59. Each case in the conformance corpus is a formula plus the answer Excel gave for it;
/// XLibur must agree, or the case must sit in <c>known-diffs.tsv</c> with a reason. The stored
/// answers come from <c>tools/excel-conformance/Record-Corpus.ps1</c>, so no Excel is needed here.
/// </summary>
public class FormulaConformanceTests
{
    [Test]
    [MethodDataSource(nameof(Cases))]
    public async Task Formula_agrees_with_Excel(Case c)
    {
        var excel = Stored(c);
        var actual = Evaluate(c);
        var matches = Matches(excel, actual);
        var known = KnownDiffs().GetValueOrDefault(c.Id);

        if (known is null)
        {
            await Assert.That(matches).IsTrue()
                .Because($"{c.Id} {c.Formula} [{c.Setup}]: Excel {excel}, XLibur {actual}. Fix it or add a reasoned row to {KnownDiffsFile}.");
            return;
        }

        await Assert.That(matches).IsFalse()
            .Because($"{c.Id} now agrees with Excel, so its {KnownDiffsFile} row is stale. Remove it.");
        await Assert.That(actual.ToString()).IsEqualTo(known.XliburAnswer)
            .Because($"{c.Id} is a known difference, but XLibur's answer changed.");
    }

    [Test]
    public async Task Known_differences_name_a_case_that_exists_and_carry_a_reason()
    {
        var ids = ReadAllCases().ToDictionary(x => x.Id);
        foreach (var diff in ReadKnownDiffs())
        {
            await Assert.That(ids.ContainsKey(diff.Id)).IsTrue().Because($"{diff.Id} is not a corpus case.");
            await Assert.That(string.IsNullOrWhiteSpace(diff.Reason)).IsFalse().Because($"{diff.Id} has no reason.");
            await Assert.That(diff.ExcelAnswer).IsEqualTo(Stored(ids[diff.Id]).ToString())
                .Because($"{diff.Id}: the stored Excel answer changed since the row was written.");
        }
    }

    [Test]
    public async Task Every_corpus_file_is_stamped_and_recorded()
    {
        // Reading throws on a missing stamp or a row without a recorded answer.
        await Assert.That(ReadAllCases().Count).IsGreaterThan(0);
    }

    [Test]
    public async Task A_corpus_file_without_an_Excel_build_stamp_is_rejected()
    {
        const string unstamped = "SUM\tmath\t\t=SUM(1)\t1\tnumber\t\n";

        await Assert.That(() => ReadCases("x.tsv", new StringReader(unstamped), []))
            .Throws<CorpusFileException>();
    }

    [Test]
    public async Task A_row_with_no_recorded_answer_is_rejected()
    {
        const string unrecorded = "# excel_build: 16.0\nSUM\tmath\t\t=SUM(1)\t\t\t\n";

        await Assert.That(() => ReadCases("x.tsv", new StringReader(unrecorded), []))
            .Throws<CorpusFileException>();
    }

    [Test]
    public async Task The_runner_reports_a_deliberately_wrong_stored_answer()
    {
        var wrong = new Case("SUM", "math", "A1=1;A2=2", "=SUM(A1:A2)", "4", "number", "", "SUM#wrong");
        var right = wrong with { ExcelValue = "3" };

        await Assert.That(Matches(Stored(wrong), Evaluate(wrong))).IsFalse();
        await Assert.That(Matches(Stored(right), Evaluate(right))).IsTrue();
    }

    [Test]
    [Arguments("number", "0.3", "number", "0.30000000000000004", true)]
    [Arguments("number", "1", "number", "1.0000000001", false)]
    [Arguments("number", "1E-310", "number", "0", true)]
    [Arguments("number", "0", "text", "0", false)]
    [Arguments("empty", "", "text", "", false)]
    [Arguments("empty", "", "number", "0", false)]
    [Arguments("text", "", "empty", "", false)]
    [Arguments("text", "a", "text", "A", false)]
    [Arguments("error", "#N/A", "error", "#VALUE!", false)]
    [Arguments("error", "#N/A", "error", "#N/A", true)]
    [Arguments("rejected", "", "error", "#NAME?", true)]
    [Arguments("rejected", "", "rejected", "", true)]
    [Arguments("rejected", "", "number", "0", false)]
    [Arguments("rejected", "", "text", "", false)]
    [Arguments("error", "#N/A", "rejected", "", false)]
    public async Task Comparison_rules(string eType, string eValue, string aType, string aValue, bool expected)
    {
        await Assert.That(Matches(new Answer(eType, eValue), new Answer(aType, aValue))).IsEqualTo(expected);
    }

    private static Dictionary<string, KnownDiff> KnownDiffs() => ReadKnownDiffs().ToDictionary(x => x.Id);

    public static IEnumerable<Func<Case>> Cases()
    {
        foreach (var c in ReadAllCases())
            yield return () => c;
    }
}
