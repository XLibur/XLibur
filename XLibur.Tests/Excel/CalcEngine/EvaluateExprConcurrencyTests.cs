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

        var threads = Enumerable.Range(0, threadCount).Select(_ => new Thread(() =>
        {
            for (var i = 0; i < rounds; i++)
            {
                barrier.SignalAndWait();
                try
                {
                    var value = XLWorkbook.EvaluateExpr(expressions[i]);
                    if (!value.IsNumber || value.GetNumber() != i * 2)
                        failures.Enqueue($"{expressions[i]} = {value}");
                }
                catch (Exception ex)
                {
                    failures.Enqueue($"{expressions[i]}: {ex.GetType().Name}: {ex.Message}");
                }
            }
        }) { IsBackground = true }).ToArray();

        foreach (var thread in threads)
            thread.Start();
        var finished = threads.All(t => t.Join(TimeSpan.FromMinutes(1)));

        await Assert.That(finished).IsTrue();
        await Assert.That(failures).IsEmpty();
    }
}
