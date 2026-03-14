using XLibur.Examples;
using NUnit.Framework;

namespace XLibur.Tests.Examples;

[TestFixture]
public class CommentsTests
{
    [Test]
    public void AddingComments()
    {
        TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
    }
}
