using NUnit.Framework;
using XLibur.Examples.Comments;

namespace XLibur.Tests.Examples;

[TestFixture]
public class CommentsTests
{
    [Test]
    [Platform("Win", Reason = "VML drawing comparison is platform-dependent: XDocument serialization produces different XML formatting on Linux vs Windows")]
    public void AddingComments()
    {
        TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
    }
}
