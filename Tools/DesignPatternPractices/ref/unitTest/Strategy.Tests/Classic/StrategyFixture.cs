using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Strategy.Classic;
namespace MarvellousWorks.PracticalPattern.Strategy.Tests.Classic
{
    [TestClass]
    public class StrategyFixture
    {
        /// <summary>
        /// 具体策略类型
        /// </summary>
        class GetMinSortStrategy : IStrategy
        {
            public int PickUp(int[] data)
            {
                if(data == null) throw new ArgumentNullException("data");
                return data.Min();
            }
        }
        class FirstDataStrategy : IStrategy
        {
            public int PickUp(int[] data)
            {
                if (data == null) throw new ArgumentNullException("data");
                return data.First();
            }
        }
        class GetMaxStrategy : IStrategy
        {
            public int PickUp(int[] data)
            {
                if (data == null) throw new ArgumentNullException("data");
                return data.Max();
            }
        }

        /// <summary>
        /// 需要采用可替换策略执行的对象
        /// </summary>
        public class Context : IContext
        {
            public IStrategy Strategy { get; set; }

            /// <summary>
            /// 执行对象依赖于策略对象的操作方法
            /// </summary>
            /// <param name="data"></param>
            /// <returns></returns>
            public int GetData(int[] data)
            {
                if (Strategy == null) throw new NullReferenceException("Strategy");
                if (data == null) throw new ArgumentNullException("data");
                return Strategy.PickUp(data);
            }
    }

        int[] data = new int[] { 5, 3, 7, 9, 1 };

        [TestMethod]
        public void Test()
        {
            var context = new Context() {Strategy = new GetMaxStrategy()};
            Assert.AreEqual<int>(context.GetData(data), 9);

            // 切换算法
            context.Strategy = new GetMinSortStrategy();
            Assert.AreEqual<int>(context.GetData(data), 1);

            // 切换算法
            context.Strategy = new FirstDataStrategy();
            Assert.AreEqual<int>(context.GetData(data), 5);
        }
    }

    
}
