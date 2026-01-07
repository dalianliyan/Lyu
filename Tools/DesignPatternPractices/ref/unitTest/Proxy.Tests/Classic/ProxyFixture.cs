using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using C = MarvellousWorks.PracticalPattern.Proxy.Classic;
namespace MarvellousWorks.PracticalPattern.Proxy.Tests.Classic
{
    [TestClass]
    public class ProxyFixture
    {
        [TestMethod]
        public void Test()
        {
            C.ISubject subject = new C.Proxy();
            Assert.AreEqual<string>("from real subject", subject.Request());
        }
    }
}
