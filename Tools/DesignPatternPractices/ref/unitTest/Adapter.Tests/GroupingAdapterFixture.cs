using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Adapter.Grouping;
namespace MarvellousWorks.PracticalPattern.Adapter.Test
{
    [TestClass]
    public class TestAdapter
    {
        [TestMethod]
        public void Test()
        {
            var factory = new FactoryMethod.Factory()
                .RegisterType<IDatabaseAdapter, OracleAdapter>("ora")
                .RegisterType<IDatabaseAdapter, SqlServerAdapter>("sql");

            var oracleAdapter = factory.Create<IDatabaseAdapter>("ora");
            var sqlAdapter = factory.Create<IDatabaseAdapter>("sql");

            // 确认之前完全不同的adaptee被统一抽象为target所定义的接口类型
            Assert.AreEqual<string>("oracle", oracleAdapter.ProviderName);
            Assert.AreEqual<string>("SQL Server", sqlAdapter.ProviderName);
        }
    }
}
