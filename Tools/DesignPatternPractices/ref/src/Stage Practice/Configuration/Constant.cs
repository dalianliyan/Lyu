namespace MarvellousWorks.PracticalPattern.StagePractice.Configuration
{
    /// <summary>
    /// 定义的常量
    /// </summary>
    public class Constant
    {
        /// <summary>
        /// 配置节组名称
        /// </summary>
        public const string SectionGroupName = "stagePractice";

        /// <summary>
        /// 认证子配置节名称
        /// </summary>
        public const string AuthenticatorSectionName = "authenticator";

        /// <summary>
        /// 认证凭证配置节名称
        /// </summary>
        public const string CredentialSectionName = "credentials";

        /// <summary>
        /// 认证非功能性控制策略配置节名称
        /// </summary>
        public const string PoliciesSectionName = "policies";

        /// <summary>
        /// 认证Provider配置节名称
        /// </summary>
        public const string ProvidersSectionName = "providers";

        /// <summary>
        /// 认证非功能性控制措施配置节名称
        /// </summary>
        public const string HandlerCoRSectionName = "handlerCoR";

        /// <summary>
        /// 认证非功能性控制措施配置元素集合名称
        /// </summary>
        public const string HandlerCollectionName = "handlers";

        /// <summary>
        /// 类型名称配置项
        /// </summary>
        public const string TypeNameItem = "value";

        /// <summary>
        /// 配置项键值
        /// </summary>
        public const string KeyItem = "key";

        /// <summary>
        /// IAuthenticationHandler CoR构造对象的配置项
        /// </summary>
        public const string BuilderItem = "builder";

        /// <summary>
        /// IAuthenticationHandler CoR中某个IAuthenticationHandler的执行次序配置项
        /// </summary>
        public const string SequenceItem = "seq";

        /// <summary>
        /// IAuthenticationHandler判定是否执行的策略名称配置项
        /// </summary>
        public const string PolicyItem = "po";

        /// <summary>
        /// 认证子类型的配置项
        /// </summary>
        public const string AuthenticatorTypeNameItem = TypeNameItem;

    }
}
