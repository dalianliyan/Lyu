using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Common;
using MarvellousWorks.PracticalPattern.Prototype.Customerized;
namespace MarvellousWorks.PracticalPattern.Prototype.Tests
{
    [TestClass]
    public class CustomerizedSerializationFixture
    {
        [TestMethod]
        public void Test()
        {
            var user = new UserInfo()
            {
                Name = "joe",
                Age = 20,
                Education = new string[] {"A", "B", "C", "D"}
            };
           
            var graph = SerializationHelper.SerializeObjectToString(user);
            var clone = SerializationHelper.DeserializeStringToObject<UserInfo>(graph);

            Assert.AreEqual<int>(3, clone.Education.Count());     // 仅序列化前三项
            CollectionAssert.AreEqual(new string[] {"A", "B", "C"}, clone.Education);

            Assert.AreNotEqual<int>(user.Age, clone.Age);      // Age没有被序列化
            Assert.AreEqual<string>(user.Name, clone.Name);    // Name被序列化
        }
    }
}
