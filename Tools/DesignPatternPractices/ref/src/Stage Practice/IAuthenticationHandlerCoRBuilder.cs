using System;
using System.Collections.Generic;
using MarvellousWorks.PracticalPattern.StagePractice.Configuration;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// 装配IAuthenticationHandler CoR的创建者
    /// </summary>
    public interface IAuthenticationHandlerCoRBuilder
    {
        /// <summary>
        /// 装配
        /// </summary>
        /// <returns>IAuthenticationHandler CoR的入口节点</returns>
        IAuthenticationHandler BuildUp(IEnumerable<IHandlerConfigurationEntry> config);

        /// <summary>
        /// 决定是否执行某个Handler的非功能性控制策略名称和类型的映射关系
        /// </summary>
        IDictionary<string, Type> PolicyRegistry { get; set; }
    }
}
