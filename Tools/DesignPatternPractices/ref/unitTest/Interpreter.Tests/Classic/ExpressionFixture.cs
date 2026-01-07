
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Interpreter.Classic;
namespace MarvellousWorks.PracticalPattern.Interpreter.Tests.Classic
{
    [TestClass]
    public class ExpressionFixture
    {
        [TestMethod]
        public void Test()
        {
            Assert.AreEqual<int>(1 + 3 - 2, new Calculator().Calculate("1+3-2"));
            Assert.AreEqual<int>(1 - 1 - 2 + 7, new Calculator().Calculate("1-1-2+7"));
        }
    }
}
