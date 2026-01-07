using System;
using System.Collections.Generic;
using System.Linq;
using MarvellousWorks.PracticalPattern.StagePractice.Configuration;
namespace MarvellousWorks.PracticalPattern.StagePractice
{
    /// <summary>
    /// <see cref="IAuthenticationHandlerCoRBuilder"/>
    /// </summary>
    public class AuthenticationHandlerCoRBuilder : IAuthenticationHandlerCoRBuilder
    {
        public IDictionary<string, Type> PolicyRegistry { get; set; }

        /// <summary>
        /// 构建Handler CoR
        /// </summary>
        /// <param name="config">登记handler的配置信息</param>
        /// <returns></returns>
        public virtual IAuthenticationHandler BuildUp(IEnumerable<IHandlerConfigurationEntry> config)
        {
            if ((config == null) || (config.Count() == 0)) return null;
            
            // 按照配置的执行次序排列非功能性控制handler
            config.OrderByDescending((x) => Convert.ToInt32(x.Sequence));
            var entries =
                (from e in config
                orderby Convert.ToInt32(e.Sequence)
                 select new
                 {
                    Handler = (IAuthenticationHandler)Activator.CreateInstance(Type.GetType(e.TypeName)),
                    Policy = (IAuthenticationPolicy)Activator.CreateInstance(PolicyRegistry[e.Policy])
                 }).ToArray();

            // 定义Handler CoR返回结果的根节点
            var root = entries.First().Handler;
            root.Policy = entries.First().Policy;
            var current = root;

            // 建立Handler CoR 的链表关系
            if (entries.Count() > 1)
            {
                for (var i = 1; i < entries.Count(); i++)
                {
                    entries[i].Handler.Policy = entries[i].Policy;
                    entries[i - 1].Handler.Successor = entries[i].Handler;
                }
                root.Successor = entries[1].Handler;
            }
            return root;
        }
    }
}
