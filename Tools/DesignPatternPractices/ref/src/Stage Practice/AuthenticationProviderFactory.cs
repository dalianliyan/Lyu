using System;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 实例化IAuthenticationProvider的工厂类型
    /// </summary>
    public class AuthenticationProviderFactory
    {
        /// <summary>
        /// 根据认证凭证类型选择合适的Provider
        /// </summary>
        /// <param name="credential"></param>
        /// <returns></returns>
        public IAuthenticationProvider Create(CredentialBase credential)
        {
            if(credential == null) throw new ArgumentNullException("credential");
            Type providerType;
            if (AuthenticationConfigurationFacade.Providers.TryGetValue(credential.GetType(), out providerType))
                return (IAuthenticationProvider)Activator.CreateInstance(providerType);
            else
                throw new NotSupportedException(credential.GetType().Name);
        }
    }
}



