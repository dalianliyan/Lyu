using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.Common.Tests
{
    [TestClass]
    public class SerializationHelperFixture
    {
        /// <summary>
        /// 测试用的可序列化类型
        /// </summary>
        [Serializable]
        class UserInfo
        {
            public string Name;
            public readonly IList<string> Education = new List<string>(); /// 引用类型

            public UserInfo GetShallowCopy() { return (UserInfo)this.MemberwiseClone(); }
            public UserInfo GetDeepCopy()
            {
                return SerializationHelper.DeserializeStringToObject<UserInfo>(SerializationHelper.SerializeObjectToString(this));
            }
        }

        /// <summary>
        /// 验证对可序列化类型进行浅表复制
        /// </summary>
        [TestMethod]
        public void ShallowCopyTest()
        {
            var user1 = new UserInfo();
            user1.Name = "joe";
            user1.Education.Add("A");
            var user2 = user1.GetShallowCopy();
            // 验证可以完成某种程度的复制
            Assert.AreEqual<string>(user1.Name, user2.Name);
            Assert.AreEqual<string>(user1.Education[0], user2.Education[0]);
            // 验证Shallow Copy方式下，没有对引用类型作出处理
            user2.Education[0] = "B";
            Assert.AreNotEqual<string>("A", user1.Education[0]);
        }

        /// <summary>
        /// 验证对可序列化类型进行深层复制
        /// </summary>
        [TestMethod]
        public void DeepCopyTest()
        {
            var user1 = new UserInfo();
            user1.Name = "joe";
            user1.Education.Add("A");
            var user2 = user1.GetDeepCopy();
            // 验证可以完成某种程度的复制
            Assert.AreEqual<string>(user1.Name, user2.Name);
            Assert.AreEqual<string>(user1.Education[0], user2.Education[0]);
            // 验证Deep Copy方式下，可以对引用类型作出处理
            user2.Education[0] = "B";
            Assert.AreEqual<string>("A", user1.Education[0]);
        }


        class Inner { }
        class Middle { Inner sub = new Inner();}
        class Outer { Middle sub = new Middle();}

        [Serializable]
        class Client
        {
            Outer outer = new Outer();
        }

        [TestMethod]
        [ExpectedException(typeof(SerializationException))]
        public void CannotSerialization()
        {
            //执行出错，会提示Outer类型没有被标明 [Serializable]
            var graph = SerializationHelper.SerializeObjectToString(new Client());
        }

        [Serializable]
        class Parent { }

        [Serializable]
        class Teacher { }

        [Serializable]
        class Child
        {
            [NonSerialized]
            public Parent[] Parents;

            public Teacher Teacher { get; set; }
        }

        [TestMethod]
        public void Test()
        {
            var child = new Child();
            child.Teacher = new Teacher();
            child.Parents = new Parent[2];
            Assert.IsNotNull(child.Parents);    // 克隆前已经非空
            Assert.IsNotNull(child.Teacher);
            var graph = SerializationHelper.SerializeObjectToString(child);
            var clone = SerializationHelper.DeserializeStringToObject<Child>(graph);
            Assert.IsNull(clone.Parents);   // 但克隆后为空
            Assert.IsNotNull(clone.Teacher);
        }
    }
}
