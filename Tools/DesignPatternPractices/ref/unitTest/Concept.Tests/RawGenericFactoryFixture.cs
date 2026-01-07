using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Concept.Generics;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    [TestClass]
    public class RawGenericFactoryFixture
    {
        interface IProduct { }
        class ConcreteProduct : IProduct { }

        interface IUser{}
        class ConcreteUserA : IUser{}

        [TestMethod]
        public void Test()
        {
            var typeName = typeof(ConcreteProduct).AssemblyQualifiedName;
            Trace.WriteLine(typeName);

            var product = new RawGenericFactory<IProduct>().Create(typeName);
            Assert.IsNotNull(product);
            Assert.IsInstanceOfType(product, typeof (ConcreteProduct));
            Assert.IsTrue(product is IProduct);

            var user = new RawGenericFactory<IUser>().Create(typeof (ConcreteUserA).AssemblyQualifiedName);
            Assert.IsNotNull(user);
            Assert.IsInstanceOfType(user, typeof(ConcreteUserA));
            Assert.IsTrue(user is IUser);
        }
    }
}
