using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;
using MarvellousWorks.PracticalPattern.StagePractice.Configuration;
using MarvellousWorks.PracticalPattern.StagePractice.Properties;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 认证过程相关配置访问对象
    /// </summary>
    /// <remarks>为了简化示例中仅支持WinForm应用</remarks>
    public static class AuthenticationConfigurationFacade
    {
        static FileSystemWatcher watcher;

        #region 简化配置访问的属性方法

        public static IDictionary<string, Type> CredentialTypes { get; private set; }
        public static IDictionary<string, Type> Policies { get; private set; }
        public static IDictionary<Type, Type> Providers { get; private set; }
        public static IAuthenticationHandler HandlerCoR { get; private set; }
        public static string AuthenticatorTypeName { get; private set; }

        #endregion

        static AuthenticationConfigurationFacade()
        {
            Load();

            // 绑定file watcher
            watcher = new FileSystemWatcher(Environment.CurrentDirectory);
            watcher.NotifyFilter = NotifyFilters.LastWrite;
            watcher.Filter = "*.config";
            watcher.Changed += (x, y) => Load();    // file watcher回调方法
            watcher.EnableRaisingEvents = true;     //启动 file watcher
        }

        /// <summary>
        /// 读取配置文件
        /// </summary>
        public static void Load()
        {
            ConfigurationManager.RefreshSection(GetFullSectionName(Constant.ProvidersSectionName));
            ConfigurationManager.RefreshSection(GetFullSectionName(Constant.CredentialSectionName));
            ConfigurationManager.RefreshSection(GetFullSectionName(Constant.PoliciesSectionName));
            ConfigurationManager.RefreshSection(GetFullSectionName(Constant.HandlerCoRSectionName));


            // 加载名称与credential类型的对应关系
            CredentialTypes = LoadDictionaryConfigSection<string>(Constant.CredentialSectionName, typeof(CredentialBase));

            // 加载认证非功能性控制策略类型列表
            Policies = LoadDictionaryConfigSection<string>(Constant.PoliciesSectionName, typeof(IAuthenticationPolicy));

            // 加载不同认证凭证所需的认证Provider类型
            Providers = LoadDictionaryConfigSection<Type>(Constant.ProvidersSectionName, typeof(IAuthenticationProvider), (x) => CredentialTypes[x]);

            // 加载认证子类型名称
            AuthenticatorTypeName = GetAuthenticatorTypeName();

            // 加载各种非功能性控制措施的配置信息
            HandlerCoR = GetHandlerCoR();
        }

        #region Helper Methods

        static IAuthenticationHandler GetHandlerCoR()
        {
            var section = (ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).GetSectionGroup(Constant.SectionGroupName) as AuthenticationConfigurationSectionGroup).HandlderCoR; ;
            var handlerBuilder = (IAuthenticationHandlerCoRBuilder)(Activator.CreateInstance(Type.GetType(section.BuilderTypeName)));
            handlerBuilder.PolicyRegistry = Policies;
            return handlerBuilder.BuildUp(section.GetHandlers());            
        }

        static string GetAuthenticatorTypeName()
        {
            var section = ConfigurationManager.GetSection(GetFullSectionName(Constant.AuthenticatorSectionName)) as IDictionary;
            if ((section == null) || (section.Count == 0) || string.IsNullOrEmpty(section[Constant.AuthenticatorTypeNameItem] as string))
                return typeof(Authenticator).AssemblyQualifiedName;
            return section[Constant.AuthenticatorTypeNameItem] as string;
        }

        static string GetFullSectionName(string sectionName)
        {
            if (string.IsNullOrEmpty(sectionName)) throw new ArgumentNullException("sectionName");
            return string.Format("{0}/{1}", Constant.SectionGroupName, sectionName);
        }

        /// <summary>
        /// 根据项目特点，读取System.Configuration.DictionarySectionHandler配置节的公共方法
        /// </summary>
        /// <typeparam name="TKey">键值类型</typeparam>
        /// <param name="sectionName">配置节名称</param>
        /// <param name="exceptedType">"value"属性定义的AssemblyQualifiedName类型名称所期望的目标类型</param>
        /// <param name="keyConverter">配置键值string类型到TKey的转换委托函数</param>
        /// <returns>转化后的配置集合信息</returns>
        static IDictionary<TKey, Type> LoadDictionaryConfigSection<TKey>(string sectionName, Type exceptedType, Func<string, TKey> keyConverter)
        {
            if(string.IsNullOrEmpty(sectionName)) throw new AbandonedMutexException("sectionName");
            if(exceptedType == null) throw new ArgumentNullException("exceptedType");

            IDictionary<TKey, Type> result = new Dictionary<TKey, Type>();
            // 读取配置文件中的 System.Configuration.DictionarySectionHandler 信息
            var config = ConvertDictionaryConfigurationSection(ConfigurationManager.GetSection(GetFullSectionName(sectionName)));
            if (config == null) throw new ConfigurationErrorsException(sectionName);
            foreach (var entry in config)
            {
                var type = Type.GetType(entry.Value);
                CheckTypeAssignment(type, exceptedType);
                if (keyConverter == null)
                    result.Add((TKey)Convert.ChangeType(entry.Key, typeof(TKey)), type);
                else
                    result.Add(keyConverter(entry.Key), type);
            }
            return result;
        }

        static IDictionary<TKey, Type> LoadDictionaryConfigSection<TKey>(string sectionName, Type exceptedType)
        {
            return LoadDictionaryConfigSection<TKey>(sectionName, exceptedType, null);
        }

        static IEnumerable<KeyValuePair<string, string>> ConvertDictionaryConfigurationSection(object target)
        {
            if (target == null) return null;
            var config = target as Hashtable;
            var result = new Dictionary<string, string>();
            foreach (DictionaryEntry entry in config)
                result.Add(entry.Key.ToString(), entry.Value.ToString());
            return result;
        }

        /// <summary>
        /// 验证类型转换
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        static void CheckTypeAssignment(Type source, Type target)
        {
            if ((source == null) || (target == null))
                throw new ArgumentNullException();
            if (!target.IsAssignableFrom(source))
                throw new InvalidCastException(
                    string.Format(Resources.ExceptionCanNotConvertType, source.FullName, target.FullName));
        }

        #endregion
    }
}
