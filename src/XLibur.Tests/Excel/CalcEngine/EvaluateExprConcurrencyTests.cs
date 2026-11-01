using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

public class EvaluateExprConcurrencyTests
{
    [Test]
    public async Task Concurrent_callers_evaluating_the_same_new_expression_each_get_its_value()
    {
        // XLWorkbook.EvaluateExpr is static, so callers that know nothing of each other share
        // whatever engine sits behind it. Each round, every thread evaluates the same fresh string
        // instance at the same moment - one the engine's parse cache has never seen.
        const int threadCount = 8;
        const int rounds = 300;
        var expressions = Enumerable.Range(0, rounds).Select(i => $"{i}*2").ToArray();
        var failures = new ConcurrentQueue<string>();
        using var barrier = new Barrier(threadCount);

        var threads = Enumerable.Range(0, threadCount)
            .Select(_ => new Thread(() => EvaluateRounds(expressions, barrier, failures)) { IsBackground = true })
            .ToArray();

        foreach (var thread in threads)
            thread.Start();
        var finished = threads.All(t => t.Join(TimeSpan.FromMinutes(1)));

        await Assert.That(finished).IsTrue();
        await Assert.That(failures).IsEmpty();
    }

    /// <summary>Evaluates each round's expression in step with the other threads, recording any wrong answer.</summary>
    private static void EvaluateRounds(string[] expressions, Barrier barrier, ConcurrentQueue<string> failures)
    {
        for (var i = 0; i < expressions.Length; i++)
        {
            barrier.SignalAndWait();
            EvaluateOne(expressions[i], i * 2, failures);
        }
    }

    private static void EvaluateOne(string expression, int expected, ConcurrentQueue<string> failures)
    {
        try
        {
            var value = XLWorkbook.EvaluateExpr(expression);
            if (!value.IsNumber || value.GetNumber() != expected)
                failures.Enqueue($"{expression} = {value}");
        }
        catch (Exception ex)
        {
            failures.Enqueue($"{expression}: {ex.GetType().Name}: {ex.Message}");
        }
    }
}
