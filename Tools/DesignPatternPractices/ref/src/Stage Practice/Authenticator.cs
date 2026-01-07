using System.Diagnostics;
using System.Security.Principal;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 默认的认证子实现
    /// </summary>
    public class Authenticator : IAuthenticator
    {
        /// <summary>
        /// 默认的认证过程
        /// </summary>
        /// <param name="credential"></param>
        /// <returns></returns>
        public virtual IIdentity Authenticate(CredentialBase credential)
        {
            Trace.WriteLine("credential type is " + credential.GetType().Name);
            // 正常的验证过程
            var result = Provider.Verify(credential);

            // 非功能性控制策略);
            if (AuthenticationConfigurationFacade.HandlerCoR != null)
                AuthenticationConfigurationFacade.HandlerCoR.Handle(credential);

            Trace.WriteLine("authentication result : " + result);
            return result;
        }

        /// <summary>
        /// 认证的实际Provider类型
        /// </summary>
        public IAuthenticationProvider Provider { get; set; }
    }
}
