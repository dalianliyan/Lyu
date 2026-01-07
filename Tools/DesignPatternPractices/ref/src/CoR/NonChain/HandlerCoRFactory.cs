using System.Collections.Generic;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.CoR.NonChain
{
    /// <summary>
    /// 构造Handler CoR的工厂类型
    /// </summary>
    public class HandlerCoRFactory
    {
        static IEnumerable<IHandler> registry = new List<IHandler>
            {
                new InternalHandler(),
                new MailPurchaseHandler(),
                new MailReturnHandler(),
                new DiscountPurchaseHandler(),
                new DiscountReturnHandler(),
                new RegularHandler()
            };

        /// <summary>
        /// 构造适于不同业务流程操作的Handler CoR
        /// </summary>
        /// <param name="option">业务流程</param>
        /// <returns>IEnumerable形式的Handler CoR</returns>
        public IEnumerable<IHandler> CreateHandlerCoR(RequestOptions option)
        {
            // 筛选适于当前业务流程的Handler
            return registry.Where(x => (x.Option & option) == option);
        }
    }
}
