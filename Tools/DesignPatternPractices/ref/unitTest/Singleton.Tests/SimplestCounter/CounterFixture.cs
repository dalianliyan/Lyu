using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Singleton.SimplestCounter;
namespace MarvellousWorks.PracticalPattern.Singleton.Tests.SimplestCounter
{
    [TestClass]
    public class CounterFixture
    {
        [TestMethod]
        public void Test()
        {
            Counter.Instance.Reset();
            Assert.AreEqual<int>(1, Counter.Instance.Next);
            Assert.AreEqual<int>(2, Counter.Instance.Next);
            Assert.AreEqual<int>(3, Counter.Instance.Next);
        }
    }
}
