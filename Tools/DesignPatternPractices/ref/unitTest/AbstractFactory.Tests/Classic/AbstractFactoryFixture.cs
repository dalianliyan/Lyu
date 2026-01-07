using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.AbstractFactory.Classic;
namespace MarvellousWorks.PracticalPattern.AbstractFactory.Tests.Classic
{
    [TestClass]
    public class TestAbstractFactory
    {
        [TestMethod]
        public void Test()
        {
            var factory = new ConcreteFactory2();
            var productA = factory.CreateProducctA();
            var productB = factory.CreateProducctB();
            Assert.IsTrue(productA is ProductA2Y);
            Assert.IsTrue(productB is ProductB2);
        }
    }
}
