using MarvellousWorks.PracticalPattern.Proxy.Tests.RemoteSubject;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.Proxy.Tests.Remote
{
    [TestClass]
    public class ProxyFixture
    {
        [TestMethod]
        public void Test()
        {
            ISubject subject = new SubjectClient();
            Assert.AreEqual<string>("from real subject", subject.Request());
        }
    }
}
