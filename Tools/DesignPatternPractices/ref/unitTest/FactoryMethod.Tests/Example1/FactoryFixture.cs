using Microsoft.VisualStudio.TestTools.UnitTesting;
using E = MarvellousWorks.PracticalPattern.FactoryMethod.Example1;
namespace MarvellousWorks.PracticalPattern.FactoryMethod.Tests.Example1
{
    [TestClass]
    public class FactoryFixture
    {
        /// <summary>
        /// 说明可以按照要求生成抽象类型，但具体实例化哪个类型由工厂决定。
        /// </summary>
        [TestMethod]
        public void Test()
        {
            Assert.IsInstanceOfType(new E.Factory().Create(), typeof (E.ConcreteProductA));
        }
    }
}
