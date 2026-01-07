using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using C = MarvellousWorks.PracticalPattern.Composite.Iterating;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Composite.Tests.Iterating
{
    [TestClass]
    public class CompositeFixture
    {
        #region non-linq version

        class LeafMatchRule : C.IMatchRule
        {
            public bool IsMatch(C.Component target)
            {
                if (target == null) return false;
                return target.GetType().IsAssignableFrom(typeof(C.Leaf));
            }
        }

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
            factory.Create<C.Leaf>(branch, "peter");                        // 11
        }

        [TestMethod]
        public void Test()
        {
            Assert.AreEqual<int>(7, corporate.Enumerate(new LeafMatchRule()).Count());
            Assert.AreEqual<int>(11, corporate.Enumerate().Count());

            //  验证通过增加遍历规则可以只显示所有leaf节点
            Trace.WriteLine("List all leaves:\n------------------------\n");
            foreach (var item in corporate.Enumerate(new LeafMatchRule()))
                Trace.WriteLine(item.Name);
        }

        #endregion

    }
}
