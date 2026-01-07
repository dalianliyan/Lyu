using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Configuration;
using System.Collections.Specialized;
namespace MarvellousWorks.PracticalPattern.FactoryMethod.Tests.Classic
{
    [TestClass]
    public class ConfigurationFactoryFixture
    {
        const string SectionName = "factoryMethod.customFactories";
        const string Name = "MarvellousWorks.PracticalPattern.FactoryMethod.Tests.Classic.IFactory, FactoryMethod.Tests";

        [TestMethod]
        public void Test()
        {
            string typeName = ((NameValueCollection)ConfigurationManager.GetSection(SectionName))[Name];
            Assert.IsTrue(typeof(IFactory).IsAssignableFrom(Type.GetType(typeName)));
        }
    }
}
