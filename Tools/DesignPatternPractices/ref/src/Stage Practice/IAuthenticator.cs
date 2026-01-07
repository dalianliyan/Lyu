using System.Security.Principal;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 认证子
    /// </summary>
    public interface IAuthenticator
    {
        /// <summary>
        /// 认证
        /// </summary>
        /// <param name="credential"></param>
        /// <returns></returns>
        IIdentity Authenticate(CredentialBase credential);

        /// <summary>
        /// 认证的实际Provider类型
        /// </summary>
        IAuthenticationProvider Provider { get; set; }
    }
}
