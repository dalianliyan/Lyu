using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.State.Raw;
namespace MarvellousWorks.PracticalPattern.State.Tests.Raw
{
    [TestClass]
    public class StoplightFixture
    {
        [TestMethod]
        public void Test()
        {
            var stopLight = new StopLight();  // green
            var values = (Color[])Enum.GetValues(typeof (Color));
            for (int i = 0; i < 100; i++ )
                Assert.AreEqual(values[i%3], stopLight.Next);
        }
    }
}
