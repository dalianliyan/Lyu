using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Singleton.SimplestCounter;
using MarvellousWorks.PracticalPattern.Common;
namespace MarvellousWorks.PracticalPattern.Singleton.Tests
{
    [TestClass]
    public class CounterFixture
    {
        const int TestTimes = 10;

        // Single thread mode
        [TestMethod]
        public void SequentialExecuteCounter()
        {
            var c1 = Counter.Instance;
            var c2 = Counter.Instance;
            Assert.IsTrue(c1 == c2);
            Assert.AreEqual<int>(c1.GetHashCode(), c2.GetHashCode());

            var counter = Counter.Instance;
            for(var i=0; i<TestTimes; i++)
                Assert.AreEqual<int>(i+1, counter.Next);
            counter.Reset();
            for (var i = 0; i < TestTimes; i++)
                Assert.AreEqual<int>(i + 1, counter.Next);
        }

        static IList<Counter> counters;

        // multiple thread mode
        [TestMethod]
        public void ParallelExecuteCounter()
        {
            counters = new List<Counter>();
            TestHarness.Parallel(
                ThreadBody,
                ThreadBody,
                ThreadBody);

            // 只有唯一实例
            Assert.AreEqual<int>(1, counters.Distinct().Count());
            // 唯一实例执行正常
            Assert.AreEqual<int>(1, Counter.Instance.Next);
        }

        void ThreadBody()
        {
            for(int i=0; i<TestTimes; i++)
            {
                Thread.Sleep(new Random().Next()%53);
                counters.Add(Counter.Instance);
            }
        }
    }
}
