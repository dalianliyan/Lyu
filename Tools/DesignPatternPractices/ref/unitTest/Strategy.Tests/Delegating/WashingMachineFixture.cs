using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Strategy.Tests.Delegating
{
    [TestClass]
    public class WashingMachineFixture
    {
        class Color { }

        class WashingMachine
        {
            public double Price { get; set; }
            public double Area { get; set; }
            public string Brand { get; set; }
            public Color Color { get; set; }
            public bool HasSmartController { get; set; }
        }

        /// <summary>
        /// 数据
        /// </summary>
        IEnumerable<WashingMachine> washingMachines;
        
        [TestInitialize]
        public void Initialize()
        {
            washingMachines = new List<WashingMachine>();
        }

        [TestMethod]
        public void TestDelegateStrategyQuery()
        {
            Predicate<WashingMachine> filter = (x) => (x.Area > 10); // 条件1
            filter += (x) => (x.Price > 200); // 条件2
            filter += (x) => (x.HasSmartController); // 条件3


            var result = washingMachines.Where(x => filter(x));
        }
    }
}
