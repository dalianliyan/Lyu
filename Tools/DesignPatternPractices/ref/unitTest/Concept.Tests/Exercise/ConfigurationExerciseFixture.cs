using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Concept.Tests.Properties;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.Exercise
{
    [TestClass]
    public class ConfigurationExerciseFixture
    {
        #region 1

        [TestMethod]
        public void TestReadSingleTagConfigurationSection()
        {
            var section = ConfigurationManager.GetSection("singleLine") as Hashtable;
            Assert.AreEqual("first", section["name"]);
            Assert.AreEqual("second", section["age"]);
            Assert.AreEqual("third", section["level"]);
        }

        [TestMethod]
        public void TestReadDictionaryConfigurationSection()
        {
            var section = ConfigurationManager.GetSection("dictionary") as Hashtable;
            Assert.AreEqual("first", section["name"]);
            Assert.AreEqual("second", section["age"]);
            Assert.AreEqual("third", section["level"]);
        }

        #endregion

        #region 2

        //为了便于访问配置内容，定义两个实体类型User和Runtime，按照”自我检验“ 的要求采用LINQ to XML方式完成示例如下：

        class User
        {
            public string Name { get; set;}
            public int Age { get; set; }
            public string Level{ get; set;}

            public override bool Equals(object obj)
            {
                var user = (User) obj;
                return string.Equals(user.Name, Name) && string.Equals(user.Level, Level) && (user.Age == Age);
            }

            public override int GetHashCode()
            {
                return base.GetHashCode() + Name.GetHashCode() + Age.GetHashCode() + Level.GetHashCode();
            }
        }

        class Runtime : User{}

        class ExerciseConfigurationManager
        {
            const string FileName = "ExerciseConfig";
            const string UsersItem = "users";
            const string UserItem = "user";
            const string RuntimeItem = "runtime";
            const string NameItem = "name";
            const string AgeItem = "age";
            const string LevelItem = "level";

            public IEnumerable<User> Users
            {
                get
                {
                    return
                        from item in XElement.Parse(Resource.ExerciseConfig).Element(UsersItem).Elements(UserItem)
                        select new User()
                                   {
                                       Name = item.Attribute(NameItem).Value,
                                       Age = Convert.ToInt32(item.Attribute(AgeItem).Value),
                                       Level = item.Attribute(LevelItem).Value
                                   };
                }
            }

            public Runtime Runtime
            {
                get
                {
                    var element = XElement.Parse(Resource.ExerciseConfig).Element(RuntimeItem);
                    return new Runtime()
                                     {
                                         Name = element.Attribute(NameItem).Value,
                                         Age = Convert.ToInt32(element.Attribute(AgeItem).Value),
                                         Level = element.Attribute(LevelItem).Value
                                     };
                }
            }

        }

        /// <summary>
        //<?xml version="1.0" encoding="utf-8" ?>
        //<config>
        //  <!--集合配置项-->
        //  <users>
        //    <user name="11" age="12" level="13"/>
        //    <user name="21" age="22" level="23"/>
        //    <user name="31" age="32" level="33"/>
        //  </users>
        //  <!--单元素配置项-->
        //  <runtime name="1" age="2" level="3"/>
        //</config>
        /// </summary>
        [TestMethod]
        public void TestLoadConfigurationWithXLinq()
        {
            var config = new ExerciseConfigurationManager();
            var users = config.Users.ToArray();
            Assert.IsNotNull(users);
            Assert.AreEqual<int>(3, users.Count());
            Assert.AreEqual<User>(users[1], new User() {Name = "21", Age = 22, Level = "23"});

            var runtime = config.Runtime;
            Assert.IsNotNull(runtime);
            Assert.AreEqual<Runtime>(runtime, new Runtime(){Name="1", Age=2, Level = "3"});
        }

        #endregion
    }
}
