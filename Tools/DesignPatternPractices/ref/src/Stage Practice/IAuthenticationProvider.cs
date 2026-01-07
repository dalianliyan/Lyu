using System.Security.Principal;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 针对具体凭证类型的认证功能的提供者
    /// </summary>
    public interface IAuthenticationProvider
    {
        /// <summary>
        /// 验证凭证信息，反馈认证结果
        /// </summary>
        /// <param name="credential"></param>
        /// <returns></returns>
        IIdentity Verify(CredentialBase credential);
    }
}
