using MarvellousWorks.PracticalPattern.Concept.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.DependencyInjection.Setter
{
    /// <summary>
    /// 通过Setter实现注入
    /// </summary>
    class Client
    {
        public ITimeProvider TimeProvider
        {
            get;
            set;
        }
    }

    [TestClass]
    public class TestClient
    {
        [TestMethod]
        public void Test()
        {
            ITimeProvider timeProvider = (new Assembler()).Create<ITimeProvider>();
            Assert.IsNotNull(timeProvider);     // 确认可以正常获得抽象类型实例
            Client client = new Client();
            client.TimeProvider = timeProvider; // 通过Setter实现注入
        }
    }
}
