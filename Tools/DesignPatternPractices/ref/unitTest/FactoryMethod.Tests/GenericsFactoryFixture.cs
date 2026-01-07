using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.FactoryMethod.Tests
{
    [TestClass]
    public class GenericsFactoryFixture
    {
        [TestMethod]
        public void CreateProductByName()
        {
            var factory = 
                new MarvellousWorks.PracticalPattern.FactoryMethod.Generics.Factory();

            // default concrete product type
            Assert.IsInstanceOfType(factory.Create(), typeof(ProductA));

            // concrete product type by name
            Assert.IsInstanceOfType(factory.Create("a"), typeof(ProductA));
            Assert.IsInstanceOfType(factory.Create("b"), typeof(ProductB));
        }

        [TestMethod]
        public void BatchCreateProductByName()
        {
            var batchSize = 5;
            var factory = new MarvellousWorks.PracticalPattern.FactoryMethod.Generics.BatchFactory();

            // default concrete product type
            var products = factory.Create(batchSize);
            Assert.AreEqual<int>(batchSize, products.Count());
            Assert.AreEqual<int>(batchSize, products.Count(x => x is ProductA));

            // concrete product type by name
            products = factory.Create("a", batchSize);
            Assert.AreEqual<int>(batchSize, products.Count());
            Assert.AreEqual<int>(batchSize, products.Count(x => x is ProductA));

            products = factory.Create("b", batchSize);
            Assert.AreEqual<int>(batchSize, products.Count());
            Assert.AreEqual<int>(batchSize, products.Count(x => x is ProductB));
        }
    }
}
