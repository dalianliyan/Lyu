using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Common.Tests
{
    [TestClass]
    public class TestHarnessFixture
    {
        IList<string> recorder = new List<string>();

        [TestMethod]
        public void TestParallelExecution()
        {
            var start = DateTime.Now;
            TestHarness.Parallel
                (
                    ()=> { Thread.Sleep(2000); recorder.Add("A");},
                    () => { Thread.Sleep(3000); recorder.Add("B"); },
                    () => { Thread.Sleep(1500); recorder.Add("C"); }
                );
            var end = DateTime.Now;

            // max execution time span
            Assert.AreEqual<int>(3, (end - start).Seconds); 

            // execution sequence
            Assert.AreEqual<string>("C", recorder[0]);
            Assert.AreEqual<string>("A", recorder[1]);
            Assert.AreEqual<string>("B", recorder[2]);
        }
    }
}
