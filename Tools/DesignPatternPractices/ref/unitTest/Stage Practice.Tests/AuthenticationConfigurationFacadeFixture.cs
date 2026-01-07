using System;
using System.Threading;
using System.Linq;
using System.Xml;
using MarvellousWorks.PracticalPattern.StagePractice.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.StagePractice.Tests.Mock;
namespace MarvellousWorks.PracticalPattern.StagePractice.Tests
{
    [TestClass]
    public class AuthenticationConfigurationFacadeFixture
    {
        const string ValueItem = "value";
        const string KeyItem = "key";
        const string ProvidersXPath = @"/configuration/stagePractice/providers/add";
        const string HandlerBuilderXPath = @"/configuration/stagePractice/handlerCoR";
        const int SleepMillSeconds = 3000;
        readonly static string ConfigFileName = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;

        #region Load configuration

        /// <summary>
        /// 确认正常读取Provider
        /// </summary>
        [TestMethod]
        public void TestLoadAuthenticationProvider()
        {
            Assert.AreEqual<Type>(AuthenticationConfigurationFacade.Providers[typeof(UsbKeyCredential)],
                typeof(UsbKeyProvider));
            Assert.AreEqual<Type>(AuthenticationConfigurationFacade.Providers[typeof(WindowsCredential)],
                typeof(WindowsProvider));
            Assert.AreEqual<Type>(AuthenticationConfigurationFacade.Providers[typeof(UserNameCredential)],
                typeof(UserNameProvider));
        }

        /// <summary>
        /// 确认正常读取Policy
        /// </summary>
        [TestMethod]
        public void TestLoadAuthenticationPolicy()
        {
            Assert.AreEqual<Type>(AuthenticationConfigurationFacade.Policies["usbPo"], typeof(UsbKeyPolicy));
            Assert.AreEqual<Type>(AuthenticationConfigurationFacade.Policies["allPo"], typeof(AllPolicy));
        }

        /// <summary>
        /// 确认正常读取Handler
        /// </summary>
        [TestMethod]
        public void TestLoadAuthenticationHandler()
        {
            var handlerRoot = AuthenticationConfigurationFacade.HandlerCoR;
            Assert.IsInstanceOfType(handlerRoot, typeof(LogClrVersionHandler));
            Assert.IsInstanceOfType(handlerRoot.Policy, typeof(AllPolicy));
            Assert.IsInstanceOfType(handlerRoot.Successor, typeof(LogOsVersionHandler));
            Assert.IsInstanceOfType(handlerRoot.Successor.Successor, typeof(LogProcessorCountHandler));
        }

        /// <summary>
        /// 确认正常读取Authenticator
        /// </summary>
        [TestMethod]
        public void TestLoadAuthenticatorType()
        {
            Assert.AreEqual<Type>(typeof(NewAuthenticator), Type.GetType(AuthenticationConfigurationFacade.AuthenticatorTypeName));
        }

        /// <summary>
        /// 确认正常读取Handler CoR Builder
        /// </summary>
        [TestMethod]
        public void TestLoadConfigHandlerCoRBuilderType()
        {
            var doc = new XmlDocument();
            doc.Load(ConfigFileName);
            var typeName =
                doc.SelectSingleNode(HandlerBuilderXPath).Attributes[Constant.BuilderItem].Value;
            Assert.AreEqual<Type>(typeof(NewHandlerCoRBuilder), Type.GetType(typeName));
            Assert.IsTrue(typeof(IAuthenticationHandlerCoRBuilder).IsAssignableFrom(typeof(NewHandlerCoRBuilder)));
        }

        #endregion

        /// <summary>
        /// 验证AuthenticationConfigurationFacade的file watcher 能否根据配置文件修改动态加载最新的配置
        /// </summary>
        /// <remarks>
        ///     为了跟踪效果，建议每次执行前重新build一下该项目，更新App.Config
        ///     另外应检查是否启动了多个UnitTest进程阻止修改.config文件
        /// </remarks>
        [TestMethod]
        public void TestConfigFileWatcher()
        {
            var credentials = AuthenticationConfigurationFacade.CredentialTypes;
            var providers = AuthenticationConfigurationFacade.Providers;
            
            // 确认原始配置信息是不同的
            var key1 = credentials.ElementAt(0).Key;
            var key2 = credentials.ElementAt(1).Key;
            Assert.AreNotEqual<Type>(providers[credentials[key1]], providers[credentials[key2]]);

            // 修改配置信息
            ModifyConfigFile(key1, key2);

            // 等待 file watcher 生效
            Thread.Sleep(SleepMillSeconds);

            // 重新获取最新的配置信息
            providers = AuthenticationConfigurationFacade.Providers;
            // 确认配置信息已经更新
            Assert.AreEqual<Type>(providers[credentials[key1]], providers[credentials[key2]]);
        }

        void ModifyConfigFile(string keySource, string keyTarget)
        {
            var doc = new XmlDocument();
            doc.Load(ConfigFileName);
            var nodes = doc.SelectNodes(ProvidersXPath);
            var nodeSource = GetNodeByKey(keySource, nodes);
            var nodeTarget = GetNodeByKey(keyTarget, nodes);
            nodeTarget.Attributes[ValueItem].Value = nodeSource.Attributes[ValueItem].Value;
            doc.Save(ConfigFileName);   // 保存对配置文件的修改
        }

        XmlNode GetNodeByKey(string key, XmlNodeList list)
        {
            if(string.IsNullOrEmpty(key)) throw new ArgumentNullException("key");
            if((list == null) || (list.Count == 0)) return null;
            for (int i = 0; i < list.Count; i++)
                if (string.Equals(list[i].Attributes[KeyItem].Value, key))
                    return list[i];
            return null;
        }

        /// <summary>
        /// 验证直接修改XML配置文件
        /// </summary>
        /// <remarks>
        ///     为了跟踪效果，建议每次执行前重新build一下该项目，更新App.Config
        ///     另外应检查是否启动了多个UnitTest进程阻止修改.config文件
        /// </remarks>
        [TestMethod]
        public void TestModifyConfiguration()
        {
            var credentials = AuthenticationConfigurationFacade.CredentialTypes;
            var doc = new XmlDocument();
            doc.Load(ConfigFileName);

            // 这里xml node的indexer为2而不是1,是因为NameValueCollection和XmlNode所设置的indexer基数不同有关系
            var node1 = doc.SelectSingleNode(ProvidersXPath + "[1]");
            var node2 = doc.SelectSingleNode(ProvidersXPath + "[2]");
            Assert.AreNotEqual<string>(node1.Attributes[ValueItem].Value, node2.Attributes[ValueItem].Value);
            
            var key1 = node1.Attributes[KeyItem].Value;
            var key2 = node2.Attributes[KeyItem].Value;

            var providers = AuthenticationConfigurationFacade.Providers;
            Assert.AreNotEqual<Type>(providers[credentials[key1]], providers[credentials[key2]]);
            var value2 = node2.Attributes[ValueItem].Value; // 原始配置信息
            node2.Attributes[ValueItem].Value = node1.Attributes[ValueItem].Value;
            doc.Save(ConfigFileName);
            Thread.Sleep(SleepMillSeconds);

            // 更新配置使两项配置相同，验证
            providers = AuthenticationConfigurationFacade.Providers;            
            Assert.AreEqual<Type>(providers[credentials[key1]], providers[credentials[key2]]);
        }
    }
}
