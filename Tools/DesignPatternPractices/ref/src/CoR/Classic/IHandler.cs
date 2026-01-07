namespace MarvellousWorks.PracticalPattern.CoR.Classic
{
    /// <summary>
    /// 抽象的操作对象
    /// </summary>
    public interface IHandler
    {
        /// <summary>
        /// 处理客户程序请求
        /// </summary>
        /// <param name="request"></param>
        void HandleRequest(Request request);

        /// <summary>
        /// 后继结点
        /// </summary>
        IHandler Successor { get; set; }

        /// <summary>
        /// 适于当前Handler处理的购买类型
        /// </summary>
        PurchaseType Type { get; }
    }
}
