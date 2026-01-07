using System;
using MarvellousWorks.PracticalPattern.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Singleton.Generic;
namespace MarvellousWorks.PracticalPattern.Singleton.Tests
{
    [TestClass]
    public class GenericSingletonFixture
    {
        class U1
        {
            private U1() { }
        }

        class U2
        {
            protected U2() { }
        }

        class U3
        {

        }

        [Serializable]
        class U4
        {
            private U4(){}
        }

        [TestMethod]
        public void TestPrivateConstructor()
        {
            var x = Singleton<U1>.Instance;
            var y = Singleton<U1>.Instance;
            Assert.IsNotNull(x);
            Assert.IsNotNull(y);
            Assert.AreEqual<U1>(x, y);
        }

        [TestMethod]
        public void TestProtectedConstructor()
        {
            var x = Singleton<U2>.Instance;
            var y = Singleton<U2>.Instance;
            Assert.IsNotNull(x);
            Assert.IsNotNull(y);
            Assert.AreEqual<U2>(x, y);
        }

        [TestMethod]
        [ExpectedException(typeof(TypeInitializationException))]
        public void TestNotSupportPublicConstructor()
        {
            var x = Singleton<U3>.Instance;
        }

        [TestMethod]
        public void TestSerializationBadEffect()
        {
            var x = Singleton<U4>.Instance;
            var y = Singleton<U4>.Instance;
            Assert.IsNotNull(x);
            Assert.IsNotNull(y);
            Assert.AreEqual(x, y);

            var z = SerializationHelper.DeserializeStringToObject<U4>(
                SerializationHelper.SerializeObjectToString(x));
            Assert.IsNotNull(z);
            Assert.AreNotEqual<U4>(x, z);
        }
    }
}