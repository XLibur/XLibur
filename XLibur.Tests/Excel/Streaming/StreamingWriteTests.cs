using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;
using XLibur.Excel.Streaming;

namespace XLibur.Tests.Excel.Streaming;

public class StreamingWriteTests
{
    #region Round-trip

    [Test]
    public async Task WritesEveryValueTypeSoItLoadsBack()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Data");
            sheet.AppendRow("text", 42.5, true);
            sheet.AppendRow(new DateTime(2026, 7, 27), new TimeSpan(1, 2, 3), XLError.DivisionByZero);
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Data");

        await Assert.That(ws.Cell("A1").Value).IsEqualTo((XLCellValue)"text");
        await Assert.That(ws.Cell("B1").Value).IsEqualTo((XLCellValue)42.5);
        await Assert.That(ws.Cell("C1").GetBoolean()).IsTrue();
        await Assert.That(ws.Cell("A2").GetDateTime()).IsEqualTo(new DateTime(2026, 7, 27));
        await Assert.That(ws.Cell("B2").GetTimeSpan()).IsEqualTo(new TimeSpan(1, 2, 3));
        await Assert.That(ws.Cell("C2").Value).IsEqualTo((XLCellValue)XLError.DivisionByZero);
    }

    [Test]
    public async Task PassesFormulaStringsThroughVerbatimWithCachedValues()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Calc");
            sheet.AppendRow(2.0, 3.0);
            using (var row = sheet.AddRow())
            {
                row.Formula("A1*B1", cachedValue: 6.0);
                // The leading '=' a user would type is accepted and stripped.
                row.Formula("=A1&\"x\"", cachedValue: "2x");
                row.Formula("A1>B1");
            }

            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Calc");

        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("A1*B1");
        await Assert.That(ws.Cell("A2").CachedValue).IsEqualTo((XLCellValue)6.0);
        await Assert.That(ws.Cell("B2").FormulaA1).IsEqualTo("A1&\"x\"");
        await Assert.That(ws.Cell("B2").CachedValue).IsEqualTo((XLCellValue)"2x");
        await Assert.That(ws.Cell("C2").FormulaA1).IsEqualTo("A1>B1");
    }

    /// <summary>
    /// A cached date is stored as a serial number just like a literal one, so it needs the same
    /// number format or it reads back as a plain number.
    /// </summary>
    [Test]
    public async Task FormulaCachedDatesAndDurationsRoundTripAsSuchTypes()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Cached");
            using (var row = sheet.AddRow())
            {
                row.Formula("TODAY()", cachedValue: new DateTime(2026, 7, 27));
                row.Formula("NOW()", cachedValue: new DateTime(2026, 7, 27, 13, 45, 0));
                row.Formula("B1-A1", cachedValue: new TimeSpan(3, 30, 0));
                row.Formula("1+1", cachedValue: 2.0);
            }

            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Cached");

        await Assert.That(ws.Cell("A1").CachedValue.Type).IsEqualTo(XLDataType.DateTime);
        await Assert.That(ws.Cell("A1").CachedValue.GetDateTime()).IsEqualTo(new DateTime(2026, 7, 27));
        await Assert.That(ws.Cell("B1").CachedValue.GetDateTime())
            .IsEqualTo(new DateTime(2026, 7, 27, 13, 45, 0));
        await Assert.That(ws.Cell("C1").CachedValue.Type).IsEqualTo(XLDataType.TimeSpan);
        await Assert.That(ws.Cell("C1").CachedValue.GetTimeSpan()).IsEqualTo(new TimeSpan(3, 30, 0));

        // A numeric cached value must not pick up a date format.
        await Assert.That(ws.Cell("D1").CachedValue.Type).IsEqualTo(XLDataType.Number);
    }

    /// <summary>
    /// A formula's computed text is the value itself, not something a user typed, so a leading
    /// apostrophe must survive rather than being absorbed into a quote-prefix style.
    /// </summary>
    [Test]
    public async Task FormulaCachedTextKeepsALeadingApostrophe()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Quoted");
            sheet.AddRow().Formula("CHAR(39)&\"x\"", cachedValue: "'x");
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("Quoted").Cell("A1").CachedValue.GetText()).IsEqualTo("'x");
    }

    [Test]
    public async Task SharedStringTableReportsReferenceAndUniqueCounts()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Counts");
            for (var i = 0; i < 5; i++)
                sheet.AppendRow("repeated", "also repeated", i);

            wb.Finish();
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sst = doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable;

        // 10 text cells referencing 2 distinct strings.
        await Assert.That(sst.Count!.Value).IsEqualTo(10U);
        await Assert.That(sst.UniqueCount!.Value).IsEqualTo(2U);
    }

    [Test]
    public async Task RoundTripsRowAndCellStyles()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var header = wb.CreateStyle();
            header.Font.Bold = true;

            var money = wb.CreateStyle();
            money.NumberFormat.Format = "#,##0.00";

            var sheet = wb.AddWorksheet("Styled");
            sheet.AppendRow(["Name", "Amount"], header);
            using (var row = sheet.AddRow())
            {
                row.Cell("Widget");
                row.Cell(1234.5, money);
            }

            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Styled");

        await Assert.That(ws.Cell("A1").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("B1").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("A2").Style.Font.Bold).IsFalse();
        await Assert.That(ws.Cell("B2").Style.NumberFormat.Format).IsEqualTo("#,##0.00");
    }

    [Test]
    public async Task DistinctStylesGetDistinctIdsInInternOrder()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Many");

            // One style per row, each differing only in font size, to check that the id handed
            // out while writing still selects the right cellXf once the styles part is written.
            for (var i = 0; i < 20; i++)
            {
                var style = wb.CreateStyle();
                style.Font.FontSize = 8 + i;
                sheet.AppendRow([i], style);
            }

            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Many");

        for (var i = 0; i < 20; i++)
            await Assert.That(ws.Cell(i + 1, 1).Style.Font.FontSize).IsEqualTo(8 + i);
    }

    [Test]
    public async Task ReusesSharedStringsForRepeatedText()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Sst");
            for (var i = 0; i < 50; i++)
                sheet.AppendRow("repeated", "also repeated");

            wb.Finish();
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var sst = doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable;
            await Assert.That(sst.Count()).IsEqualTo(2);
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("Sst").Cell("A50").GetString()).IsEqualTo("repeated");
    }

    [Test]
    public async Task InlineModeWritesNoSharedStringPart()
    {
        using var ms = new MemoryStream();
        var options = new XLStreamingOptions { StringStorage = XLStreamingStringStorage.Inline };
        using (var wb = XLStreamingWorkbook.Create(ms, options))
        {
            var sheet = wb.AddWorksheet("Inline");
            sheet.AppendRow("alpha", "beta");
            sheet.AppendRow("alpha", " leading space");
            wb.Finish();
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
            await Assert.That(doc.WorkbookPart!.SharedStringTablePart).IsNull();

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Inline");
        await Assert.That(ws.Cell("A2").GetString()).IsEqualTo("alpha");
        await Assert.That(ws.Cell("B2").GetString()).IsEqualTo(" leading space");
    }

    [Test]
    public async Task SkippedRowsAndCellsLeaveGapsAtTheRightAddresses()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Gaps");
            sheet.AppendRow("first");
            sheet.SkipRows(3);
            using (var row = sheet.AddRow())
            {
                row.Cell("a");
                row.Skip(2);
                row.Cell("d");
                row.At(10).Cell("j");
            }

            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Gaps");

        await Assert.That(ws.Cell("A1").GetString()).IsEqualTo("first");
        await Assert.That(ws.Cell("A5").GetString()).IsEqualTo("a");
        await Assert.That(ws.Cell("B5").IsEmpty()).IsTrue();
        await Assert.That(ws.Cell("D5").GetString()).IsEqualTo("d");
        await Assert.That(ws.Cell("J5").GetString()).IsEqualTo("j");
    }

    [Test]
    public async Task WritesMultipleWorksheetsInOrder()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            wb.AddWorksheet("First").AppendRow("one");
            // The first sheet is completed implicitly by adding the second.
            wb.AddWorksheet("Second").AppendRow("two");
            wb.AddWorksheet("Third").AppendRow("three");
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);

        await Assert.That(loaded.Worksheets.Count).IsEqualTo(3);
        await Assert.That(loaded.Worksheets.Select(w => w.Name).ToArray())
            .IsEquivalentTo(new[] { "First", "Second", "Third" });
        await Assert.That(loaded.Worksheet("Second").Cell("A1").GetString()).IsEqualTo("two");
    }

    [Test]
    public async Task WritesAnEmptyWorksheet()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            wb.AddWorksheet("Empty");
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("Empty").LastCellUsed()).IsNull();
    }

    [Test]
    public async Task WritesToAFilePath()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlibur-streaming-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = XLStreamingWorkbook.Create(path))
            {
                wb.AddWorksheet("Sheet1").AppendRow("on disk");
                wb.Finish();
            }

            using var loaded = new XLWorkbook(path);
            await Assert.That(loaded.Worksheet("Sheet1").Cell("A1").GetString()).IsEqualTo("on disk");
        }
        finally
        {
            File.Delete(path);
        }
    }

    #endregion Round-trip

    #region Sheet-level features

    [Test]
    public async Task RoundTripsColumnWidthsFreezePanesAndAutoFilter()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Layout");
            sheet.Column(1).Width = 30;
            sheet.Columns(2, 3).Hidden = true;
            sheet.FreezeRows(1);
            sheet.AutoFilterRange = "A1:C1";

            sheet.AppendRow("Name", "Qty", "Note");
            sheet.AppendRow("Widget", 1, "ok");
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Layout");

        await Assert.That(ws.Column(1).Width).IsEqualTo(30).Within(0.01);
        await Assert.That(ws.Column(2).IsHidden).IsTrue();
        await Assert.That(ws.Column(3).IsHidden).IsTrue();
        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(1);
        await Assert.That(ws.AutoFilter.IsEnabled).IsTrue();
    }

    [Test]
    public async Task RoundTripsRowHeightAndHiddenRows()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Rows");
            sheet.AddRow(null, height: 40).Cell("tall");
            sheet.AddRow(null, hidden: true).Cell("hidden");
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws = loaded.Worksheet("Rows");

        await Assert.That(ws.Row(1).Height).IsEqualTo(40).Within(0.01);
        await Assert.That(ws.Row(2).IsHidden).IsTrue();
    }

    [Test]
    public async Task FreezingBothAxesSetsBothSplits()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("Panes");
            sheet.FreezePanes(2, 1);
            sheet.AppendRow("a", "b");
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var view = loaded.Worksheet("Panes").SheetView;

        await Assert.That(view.SplitRow).IsEqualTo(2);
        await Assert.That(view.SplitColumn).IsEqualTo(1);
    }

    #endregion Sheet-level features

    #region Package validity

    [Test]
    public async Task ProducesAPackageThatPassesOpenXmlValidation()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var bold = wb.CreateStyle();
            bold.Font.Bold = true;

            var sheet = wb.AddWorksheet("Valid");
            sheet.Column(1).Width = 25;
            sheet.FreezeRows(1);
            sheet.AutoFilterRange = "A1:C1";
            sheet.AppendRow(["Name", "Qty", "When"], bold);
            sheet.AppendRow("Widget", 3, new DateTime(2026, 1, 2));
            using (var row = sheet.AddRow())
            {
                row.Cell("Total");
                row.Formula("SUM(B2:B2)", cachedValue: 3.0);
            }

            wb.Finish();
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc).ToArray();
        var message = string.Join("\r\n",
            errors.Select(e => $"Part {e.Part?.Uri}, Path {e.Path?.XPath}: {e.Description}"));

        await Assert.That(errors).IsEmpty().Because(message);
    }

    [Test]
    public async Task ValidatesAnInlineStringPackage()
    {
        using var ms = new MemoryStream();
        var options = new XLStreamingOptions { StringStorage = XLStreamingStringStorage.Inline };
        using (var wb = XLStreamingWorkbook.Create(ms, options))
        {
            wb.AddWorksheet("Inline").AppendRow("alpha", 1, "beta");
            wb.Finish();
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc).ToArray();

        await Assert.That(errors).IsEmpty()
            .Because(string.Join("\r\n", errors.Select(e => e.Description)));
    }

    #endregion Package validity

    #region Memory

    /// <summary>
    /// The point of the whole API: resident memory must not grow with the number of rows.
    /// </summary>
    /// <remarks>
    /// Runs at 100K x 10 rather than the spec's 1M x 10 so it stays a few seconds of CI time -
    /// the 1M measurement lives in <c>StreamingWriteBenchmarks</c>. Scale is not what the test
    /// proves anyway: what it proves is that the live heap after writing is flat, which fails
    /// just as loudly at 100K if any per-row state is being retained. The equivalent
    /// <see cref="XLWorkbook"/> would hold hundreds of MB of slices at this size.
    ///
    /// The bound is deliberately far above what is measured - retained growth is on the order of
    /// tens of KB - so that JIT warmup on a cold CI agent cannot trip it, while still catching
    /// the regression this guards against. Routing the package back through
    /// <c>System.IO.Packaging</c>, which buffers every part uncompressed until close, put ~48 MB
    /// on the heap at this size.
    ///
    /// Written to a temp file, not a MemoryStream, so the package itself is not counted as heap.
    /// </remarks>
    [Test]
    public async Task StreamingAMillionCellsDoesNotGrowTheHeap()
    {
        const int rowCount = 100_000;
        const int columnCount = 10;

        var path = Path.Combine(Path.GetTempPath(), $"xlibur-streaming-mem-{Guid.NewGuid():N}.xlsx");
        try
        {
            var baseline = GC.GetTotalMemory(forceFullCollection: true);
            long liveAfterWrite;

            using (var wb = XLStreamingWorkbook.Create(path))
            {
                var sheet = wb.AddWorksheet("Big");
                var values = new XLCellValue[columnCount];

                for (var r = 0; r < rowCount; r++)
                {
                    // A bounded set of distinct strings: an unbounded one would legitimately
                    // grow the shared string dictionary, which is documented behaviour rather
                    // than a leak.
                    values[0] = $"row-kind-{r % 100}";
                    for (var c = 1; c < columnCount; c++)
                        values[c] = r * columnCount + c;

                    sheet.AppendRow(values, null);
                }

                liveAfterWrite = GC.GetTotalMemory(forceFullCollection: true);
                wb.Finish();
            }

            var growth = liveAfterWrite - baseline;
            await Assert.That(growth).IsLessThan(16L * 1024 * 1024)
                .Because($"streaming {rowCount:N0} x {columnCount} cells retained {growth / (1024 * 1024)} MB");

            // ...and the file it produced is real.
            using var loaded = new XLWorkbook(path);
            var ws = loaded.Worksheet("Big");
            await Assert.That(ws.Cell(rowCount, 1).GetString()).IsEqualTo($"row-kind-{(rowCount - 1) % 100}");
            await Assert.That(ws.Cell(rowCount, columnCount).GetDouble())
                .IsEqualTo((rowCount - 1) * columnCount + columnCount - 1);
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Owning the zip means the output is written strictly forwards, so it does not have to be
    /// seekable - a workbook can go straight to a network stream. <see cref="XLWorkbook.SaveAs(Stream)"/>
    /// cannot do this.
    /// </summary>
    [Test]
    public async Task WritesToANonSeekableStream()
    {
        using var backing = new MemoryStream();
        using (var forwardOnly = new ForwardOnlyStream(backing))
        using (var wb = XLStreamingWorkbook.Create(forwardOnly))
        {
            wb.AddWorksheet("Piped").AppendRow("no seeking", 1);
            wb.Finish();
        }

        backing.Position = 0;
        using var loaded = new XLWorkbook(backing);
        await Assert.That(loaded.Worksheet("Piped").Cell("A1").GetString()).IsEqualTo("no seeking");
    }

    [Test]
    public async Task FastestCompressionProducesALargerFileThanOptimal()
    {
        var optimal = WriteSized(CompressionLevel.Optimal);
        var fastest = WriteSized(CompressionLevel.Fastest);

        await Assert.That(fastest).IsGreaterThan(optimal);

        static long WriteSized(CompressionLevel level)
        {
            using var ms = new MemoryStream();
            using (var wb = XLStreamingWorkbook.Create(ms, new XLStreamingOptions { CompressionLevel = level }))
            {
                var sheet = wb.AddWorksheet("Data");
                for (var r = 0; r < 5_000; r++)
                    sheet.AppendRow($"row {r}", r, r * 1.5, $"note for row {r}");

                wb.Finish();
            }

            return ms.Length;
        }
    }

    /// <summary>
    /// The same knob on the ordinary save path, which reaches it through the SDK's
    /// <c>OpenXmlPackage.CompressionOption</c> rather than by owning the zip.
    /// </summary>
    [Test]
    public async Task SaveOptionsCompressionLevelChangesPackageSize()
    {
        var optimal = SaveSized(CompressionLevel.Optimal);
        var none = SaveSized(CompressionLevel.NoCompression);

        await Assert.That(none).IsGreaterThan(optimal);

        static long SaveSized(CompressionLevel level)
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                for (var r = 1; r <= 2_000; r++)
                {
                    ws.Cell(r, 1).Value = $"row {r}";
                    ws.Cell(r, 2).Value = r;
                    ws.Cell(r, 3).Value = $"note for row {r}";
                }

                wb.SaveAs(ms, new SaveOptions { CompressionLevel = level });
            }

            return ms.Length;
        }
    }

    /// <summary>A write-only, non-seekable wrapper, standing in for a network stream.</summary>
    private sealed class ForwardOnlyStream(Stream inner) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => inner.Flush();
        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
        public override void Write(ReadOnlySpan<byte> buffer) => inner.Write(buffer);
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }

    #endregion Memory

    #region Misuse

    [Test]
    public async Task RejectsColumnConfigurationAfterTheFirstRow()
    {
        using var ms = new MemoryStream();
        using var wb = XLStreamingWorkbook.Create(ms);
        var sheet = wb.AddWorksheet("Late");
        sheet.AppendRow("a");

        await Assert.That(() => sheet.Column(1).Width = 10).Throws<InvalidOperationException>();
        await Assert.That(() => sheet.FreezeRows(1)).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task RejectsWritingToARowThatIsNoLongerOpen()
    {
        using var ms = new MemoryStream();
        using var wb = XLStreamingWorkbook.Create(ms);
        var sheet = wb.AddWorksheet("Stale");

        var first = sheet.AddRow();
        first.Cell("ok");
        sheet.AddRow().Cell("second");

        // A ref struct cannot be captured in a lambda, so the assertion is written out.
        Exception caught = null;
        try
        {
            first.Cell("too late");
        }
        catch (InvalidOperationException e)
        {
            caught = e;
        }

        await Assert.That(caught).IsNotNull();
    }

    [Test]
    public async Task RejectsDuplicateSheetNames()
    {
        using var ms = new MemoryStream();
        using var wb = XLStreamingWorkbook.Create(ms);
        wb.AddWorksheet("Data");

        await Assert.That(() => wb.AddWorksheet("data")).Throws<ArgumentException>();
    }

    [Test]
    public async Task RejectsAddingWorksheetsAfterFinish()
    {
        using var ms = new MemoryStream();
        using var wb = XLStreamingWorkbook.Create(ms);
        wb.AddWorksheet("Data").AppendRow("a");
        wb.Finish();

        await Assert.That(() => wb.AddWorksheet("More")).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task FinishIsIdempotent()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            wb.AddWorksheet("Data").AppendRow("a");
            wb.Finish();
            wb.Finish();
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet("Data").Cell("A1").GetString()).IsEqualTo("a");
    }

    [Test]
    public async Task RejectsWritingCellsToTheLeftOfThePosition()
    {
        using var ms = new MemoryStream();
        using var wb = XLStreamingWorkbook.Create(ms);
        var sheet = wb.AddWorksheet("Back");
        var row = sheet.AddRow();
        row.Cell("a").Cell("b").Cell("c");

        Exception caught = null;
        try
        {
            row.At(2);
        }
        catch (ArgumentOutOfRangeException e)
        {
            caught = e;
        }

        await Assert.That(caught).IsNotNull();
    }

    #endregion Misuse
}
