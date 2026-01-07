using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Prototype.Refresh;
namespace MarvellousWorks.PracticalPattern.Prototype.Tests.Refresh
{
    [TestClass]
    public class PrototypeFixture
    {
        [TestMethod]
        public void Test()
        {
            var p1 = new ConcretePrototype();
            p1.Name = "Hello";
            var p2 = (ConcretePrototype)p1.Clone();
            Assert.AreEqual<string>("Hello", p2.Name);
        }
    }
}
