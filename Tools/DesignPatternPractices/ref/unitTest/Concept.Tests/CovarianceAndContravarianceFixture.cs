using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    [TestClass]
    public class CovarianceAndContravarianceFixture
    {
        interface II{}
        abstract class Base : II{}
        class Level1 : Base{}
        class Level2 : Level1{}


        [TestMethod]
        public void TestCovariance()
        {
            Func<II> func = () => new Level1();
            Assert.IsInstanceOfType(func(), typeof(II));
            Assert.IsInstanceOfType(func(), typeof(Level1));
        }

        [TestMethod]
        public void TestContravariance()
        {
            Action<Level1> l1Action = (x) => { };
            Action<Level2> l2Action = l1Action;
        }
    }
}
