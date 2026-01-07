using System.Xml;
using MarvellousWorks.PracticalPattern.Common;
namespace MarvellousWorks.PracticalPattern.Adapter.XML
{
    
    /// <summary>
    /// 基于XSLT的通用适配器类型
    /// </summary>
    public abstract class XsltAdapterBase
    {
        protected string xslt;
        public virtual XmlDocument Transform(XmlDocument source)
        {
            return XmlHelper.Transform(xslt, source);
        }
    }

    /// <summary>
    /// 一个的采用XSLT方式适配的示例适配器类型
    /// </summary>
    public class OrderAdapter : XsltAdapterBase
    {
        public OrderAdapter() { base.xslt = "Order.xslt"; }
    }
}
