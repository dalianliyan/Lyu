using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Builder.Reversable;
namespace MarvellousWorks.PracticalPattern.Builder.Tests
{
    [TestClass]
    public class TestBuilder
    {
        [TestMethod]
        public void Test()
        {
            IBuilder<Product> builder = new ProductBuilder();
            var product = builder.BuildUp();
            Assert.AreEqual<int>(5, product.Count);
            Assert.AreEqual<int>(5, product.Items.Count);
            product = builder.TearDown();
            Assert.AreEqual<int>(0, product.Count);
            Assert.AreEqual<int>(0, product.Items.Count);
        }
    }
}
