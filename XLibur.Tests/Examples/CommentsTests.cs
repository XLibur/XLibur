using XLibur.Examples;
using NUnit.Framework;
using XLibur.Examples.Comments;

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
