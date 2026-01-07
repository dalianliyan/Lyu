using System;
using System.Collections.Generic;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.Concept.FluentInterface
{

    /// <summary>
    /// 连贯接口的实现类
    /// </summary>
    public class CollectionDataFacade : IDataFacade
    {
        IList<Currency> data = new List<Currency>();

        public IEnumerable<Currency> ExecuteQuery(Func<Currency, bool> filter)
        {
            return filter == null ? data : data.Where(filter);
        }

        public IDataFacade AddNewCurrency(string code, string name)
        {
            if (string.IsNullOrEmpty(code)) throw new ArgumentNullException("code");
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
            data.Add(new Currency() {Code = code, Name = name});
            return this;
        }
    }
}
