using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 认证过程非功能性控制策略对象
    /// </summary>
    public interface IAuthenticationPolicy
    {
        /// <summary>
        /// 策略是否适用
        /// </summary>
        /// <param name="credential"></param>
        /// <returns></returns>
        bool IsMatch(CredentialBase credential);
    }
}
