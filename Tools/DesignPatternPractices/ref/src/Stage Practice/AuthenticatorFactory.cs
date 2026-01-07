using System;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 认证子工厂类型
    /// </summary>
    public class AuthenticatorFactory
    {
        /// <summary>
        /// 构造认证子类型实例
        /// </summary>
        /// <param name="credential"></param>
        /// <returns></returns>
        public IAuthenticator Create(CredentialBase credential)
        {
            if(credential == null) throw new ArgumentNullException("credential");
            var result = (IAuthenticator)Activator.CreateInstance(Type.GetType(AuthenticationConfigurationFacade.AuthenticatorTypeName));
            result.Provider = new AuthenticationProviderFactory().Create(credential);
            return result;
        }
    }
}
