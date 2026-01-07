using System.Collections;
using System.Configuration;
namespace MarvellousWorks.PracticalPattern.StagePractice.Configuration
{
    /// <summary>
    /// 认证机制的配置节组
    /// </summary>
    public class AuthenticationConfigurationSectionGroup : ConfigurationSectionGroup
    {
        /// <summary>
        /// 获得认证子类型名称
        /// </summary>
        /// <remarks>默认为内置的Authenticator类型</remarks>
        [ConfigurationProperty(Constant.AuthenticatorSectionName)]
        public string AuthenticatorTypeName
        {
            get
            {
                var section = base.Sections[Constant.AuthenticatorSectionName] as IDictionary;
                if ((section == null) || (section.Count == 0) || string.IsNullOrEmpty(section[Constant.AuthenticatorTypeNameItem] as string))
                    return typeof (Authenticator).AssemblyQualifiedName;
                return section[Constant.AuthenticatorTypeNameItem] as string;
            }
        }


        [ConfigurationProperty(Constant.CredentialSectionName, IsRequired = true)]
        public IDictionary Credentials
        {
            get
            {
                return base.Sections[Constant.CredentialSectionName] as IDictionary;
            }
        }

        [ConfigurationProperty(Constant.PoliciesSectionName)]
        public IDictionary Policies
        {
            get
            {
                return base.Sections[Constant.PoliciesSectionName] as IDictionary;
            }
        }

        [ConfigurationProperty(Constant.CredentialSectionName, IsRequired = true)]
        public IDictionary Providers
        {
            get
            {
                return base.Sections[Constant.CredentialSectionName] as IDictionary;
            }
        }

        [ConfigurationProperty(Constant.HandlerCoRSectionName, IsRequired = false)]
        public HandlerConfigurationSection HandlderCoR
        {
            get
            {
                return base.Sections[Constant.HandlerCoRSectionName] as HandlerConfigurationSection;
            }
        }
    }
}
