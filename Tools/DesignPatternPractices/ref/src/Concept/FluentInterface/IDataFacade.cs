using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Concept.FluentInterface
{
    public class Currency
    {
        /// <summary>
        /// 货币代码
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 货币名称
        /// </summary>
        public string Name { get; set; }
    }

    public interface IDataFacade
    {
        /// <summary>
        /// 根据筛选条件执行查询
        /// </summary>
        /// <param name="filter">筛选条件委托</param>
        /// <returns>查询结果</returns>
        IEnumerable<Currency> ExecuteQuery(Func<Currency, bool> filter);

        /// <summary>
        /// 追加新的货币记录
        /// </summary>
        /// <param name="code">货币代码</param>
        /// <param name="name">货币名称</param>
        /// <returns>连贯对象实例</returns>
        /// <remarks>连贯接口方法</remarks>
        IDataFacade AddNewCurrency(string code, string name);
    }
}
