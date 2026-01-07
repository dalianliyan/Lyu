using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using C = MarvellousWorks.PracticalPattern.Composite.Classic;
namespace MarvellousWorks.PracticalPattern.Composite.Tests.Classic
{
    [TestClass]
    public class CompositeFixture
    {
        C.Component corporate;

        /// <summary>
        /// 建立测试公司的组织结构
        /// </summary>
        [TestInitialize]
        public void Initialize()
        {
            var factory = new C.ComponentFactory();
            corporate = factory.Create<C.Composite>("corporate");           // 1
            factory.Create<C.Leaf>(corporate, "president");                 // 2
            factory.Create<C.Leaf>(corporate, "vice president");            // 3
            var sales = factory.Create<C.Composite>(corporate, "sales");    // 4
            var market = factory.Create<C.Composite>(corporate, "market");  // 5
            factory.Create<C.Leaf>(sales, "joe");                           // 6
            factory.Create<C.Leaf>(sales, "bob");                           // 7
            factory.Create<C.Leaf>(market, "judi");                         // 8
            var branch = factory.Create<C.Composite>(corporate, "branch");  // 9
            factory.Create<C.Leaf>(branch, "manager");                      // 10
            factory.Create<C.Leaf>(branch, "peter");                        // 11;
        }

        [TestMethod]
        public void Test()
        {
            // 验证确实可以把所有节点的名称（含子孙节点）遍历出来
            Assert.AreEqual<int>(11, corporate.GetNameList().Count());

            Trace.WriteLine("List all components:\n------------------------\n");
            corporate.GetNameList().ToList().ForEach(x=>Trace.WriteLine(x));
        }
    }
}

