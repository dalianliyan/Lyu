using System;
using System.Collections.Generic;
using System.Linq;

using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using MarvellousWorks.PracticalPattern.Common;

namespace MarvellousWorks.PracticalPattern.Common.Tests
{
    [TestClass]
    public class EnumerableHelperFixture
    {
        [TestMethod]
        public void TestForEach()
        {
            string[] data = { "A", "B", "CDE", "FG"};
            int[] lens = new int[data.Length];
            int i = 0;
            EnumerableHelper.ForEach<string>(data,
                delegate(string item) 
                {
                    lens[i++] = item.Length;
                    Trace.WriteLine(lens[i - 1]);
                });
            Assert.AreEqual<int>(string.Join("", data).Length, lens.Sum());
        }
    }
}
