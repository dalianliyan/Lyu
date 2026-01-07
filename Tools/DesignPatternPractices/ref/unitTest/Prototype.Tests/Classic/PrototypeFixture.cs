using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Prototype.Classic;
namespace MarvellousWorks.PracticalPattern.Prototype.Tests.Classic
{
    [TestClass]
    public class PrototypeFixture
    {
        IPrototype sample;

        [TestInitialize]
        public void Initialize()
        {
            sample = new ConcretePrototype();
        }

        [TestMethod]
        public void Test()
        {
            var image = sample.Clone();
            Assert.IsNull(sample.Name);                 //  副本与样本当时的内容一致
            Assert.IsNull(image.Name);

            sample.Name = "A";                          //  修改样本内容
            image = sample.Clone();
            Assert.AreEqual<string>("A", image.Name);   // 副本与样本当时的内容一致
            Assert.IsInstanceOfType(image, typeof(ConcretePrototype));
            image.Name = "B";                           // 独立修改副本的内容
            Assert.IsTrue(sample.Name != image.Name);   // 证明副本与样本两个独立的个体
        }

        [TestMethod]
        public void TestReferenceType()
        {
            var image = sample.Clone();
            Assert.AreEqual<int>(image.Signal.GetHashCode(), sample.Signal.GetHashCode());
        }
    }
}
