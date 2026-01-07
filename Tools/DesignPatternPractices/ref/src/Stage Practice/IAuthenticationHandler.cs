using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 认证过程非功能性控制处理对象
    /// </summary>
    public interface IAuthenticationHandler
    {
        /// <summary>
        /// 适用策略
        /// </summary>
        IAuthenticationPolicy Policy { get; set; }

        /// <summary>
        /// 处理
        /// </summary>
        /// <param name="credential"></param>
        void Handle(CredentialBase credential);

        /// <summary>
        /// 后继节点
        /// </summary>
        IAuthenticationHandler Successor { get; set; }

    }
}
