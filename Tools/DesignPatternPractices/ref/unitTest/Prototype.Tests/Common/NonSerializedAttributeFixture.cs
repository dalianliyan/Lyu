using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Common;
namespace MarvellousWorks.PracticalPattern.Prototype.Tests.Common
{
    [TestClass]
    public class NonSerializedAttributeFixture
    {
        [Serializable]
        class UserInfo
        {
            [NonSerialized]
            public string[] Parents;
        }

        [TestMethod]
        public void Test()
        {
            var user1 = new UserInfo();
            user1.Parents = new string[2];
            Assert.IsNotNull(user1.Parents);    // 克隆前已经非空
            var user2 = SerializationHelper.DeserializeStringToObject<UserInfo>(SerializationHelper.SerializeObjectToString(user1));
            Assert.IsNull(user2.Parents);       // 但克隆后为空
        }
    }
}
