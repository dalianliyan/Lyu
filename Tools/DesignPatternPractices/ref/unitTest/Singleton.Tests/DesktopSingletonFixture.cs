using System;
using System.Collections.Generic;
using System.Linq;
using MarvellousWorks.PracticalPattern.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.Singleton.Tests
{
    [TestClass]
    public class DesktopSingletonFixture
    {
        public class ThreadScopedSingleton
        {
            ThreadScopedSingleton() { }
            [ThreadStatic]  // 说明每个Instance仅在当前线程内为静态
            static ThreadScopedSingleton instance;

            public static ThreadScopedSingleton Instance
            {
                get
                {
                    if (instance == null)
                        instance = new ThreadScopedSingleton();
                    return instance;
                }
            }
        }

        static IList<ThreadScopedSingleton> registry;
        const int TestTimes = 10;
        [TestMethod]
        public void Test()
        {
            registry = new List<ThreadScopedSingleton>();
            TestHarness.Parallel(
                ThreadBody,
                ThreadBody,
                ThreadBody);
            // not singleton cross threads
            Assert.AreEqual<int>(3, registry.Distinct().Count());
        }

        void ThreadBody()
        {
            ThreadScopedSingleton instance = ThreadScopedSingleton.Instance;
            // singleton in current thread
            for(int i=0; i<TestTimes; i++)
                Assert.IsTrue(instance == ThreadScopedSingleton.Instance);
            registry.Add(instance);
        }
    }
}
