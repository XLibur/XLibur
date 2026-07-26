using XLibur.Examples;
using System.IO;
using XLibur.Tests.Utils;
using System.Threading.Tasks;

namespace XLibur.Tests;

public class ExcelDocsComparerTests
{
    [Test]
    public async Task CheckEqual()
    {
        var left = ExampleHelper.GetTempFilePath("left.xlsx");
        var right = ExampleHelper.GetTempFilePath("right.xlsx");
        try
        {
            new BasicTable().Create(left);
            new BasicTable().Create(right);
            await Assert.That(ExcelDocsComparer.Compare(left, right, out var message)).IsTrue();
        }
        finally
        {
            if (File.Exists(left))
            {
                File.Delete(left);
            }
            if (File.Exists(right))
            {
                File.Delete(right);
            }
        }
    }

    [Test]
    public async Task CheckNonEqual()
    {
        var left = ExampleHelper.GetTempFilePath("left.xlsx");
        var right = ExampleHelper.GetTempFilePath("right.xlsx");
        try
        {
            new BasicTable().Create(left);
            HelloWorld.Create(right);

            await Assert.That(ExcelDocsComparer.Compare(left, right, out var message)).IsFalse();
        }
        finally
        {
            if (File.Exists(left))
            {
                File.Delete(left);
            }
            if (File.Exists(right))
            {
                File.Delete(right);
            }
        }
    }
}
