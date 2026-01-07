namespace MarvellousWorks.PracticalPattern.CoR.NonChain
{
    /// <summary>
    /// 抽象的操作对象
    /// </summary>
    /// <remarks>因为通过LINQ方式组织链表，所以不需定义后继节点(Successor)</remarks>
    public interface IHandler
    {
        /// <summary>
        /// 处理客户程序请求
        /// </summary>
        /// <param name="request"></param>
        void HandleRequest(Request request);

        /// <summary>
        /// 适于当前Handler处理的购买类型
        /// </summary>
        PurchaseType Type { get; }

        /// <summary>
        /// 当前Handler适用的业务流程
        /// </summary>
        RequestOptions Option { get; }
    }
}
