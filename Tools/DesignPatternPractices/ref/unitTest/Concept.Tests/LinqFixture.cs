using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    [TestClass]
    public class LinqFixture
    {
       
        //  登记抽象类型与不同具体类型的映射关系
        //  登记信息具有层次性
        IDictionary<Type, IEnumerable<KeyValuePair<string, Type>>> registry =
            new Dictionary<Type, IEnumerable<KeyValuePair<string, Type>>>();

        [TestInitialize]
        public void Initialize()
        {
            registry.Add(
                    typeof (DbProviderFactory),
                    new KeyValuePair<string, Type>[]
                        {
                            new KeyValuePair<string, Type>("sql", typeof (SqlClientFactory)),
                            new KeyValuePair<string, Type>("oledb", typeof (OleDbFactory)),
                            new KeyValuePair<string, Type>("odbc", typeof (OdbcFactory)),
                        });
            registry.Add(
                    typeof(ICollection<string>),
                    new KeyValuePair<string, Type>[]
                        {
                            new KeyValuePair<string, Type>("list", typeof (List<string>)),
                            new KeyValuePair<string, Type>("hashSet", typeof (HashSet<string>)),
                            new KeyValuePair<string, Type>("sortedSet", typeof (SortedSet<string>))
                        });
        }

        // 寻找适合odbc的db provider factory
        const string TestName = "odbc";

        [TestMethod]
        public void NoLinq()
        {
            IEnumerable<KeyValuePair<string, Type>> dbFactories;
            if (!registry.TryGetValue(typeof (DbProviderFactory), out dbFactories))
            {
                return; // 没找到DbProviderFactory类型相关的注册项
            }

            Type targetType = null;
            foreach (var item in dbFactories)
            {
                if (string.Equals(TestName, item.Key))
                {
                    targetType = item.Value;
                    break;
                }
            }

            Assert.IsNotNull(targetType);
            // 确认找到odbc数据库对应的工厂类型
            Assert.AreEqual<Type>(typeof(OdbcFactory), targetType);
        }

        [TestMethod]
        public void WithLinq()
        {
            Assert.AreEqual<Type>(
                registry.FirstOrDefault(x => x.Key == typeof (DbProviderFactory)).Value.
                FirstOrDefault(x => x.Key.Equals(TestName)).Value, 
                typeof(OdbcFactory));
        }
    }
}
