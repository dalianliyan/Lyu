namespace MarvellousWorks.PracticalPattern.CoR.Classic
{
    #region 具体操作类型

    /// <summary>
    /// 只适于“内部价格”的操作
    /// </summary>
    public class InternalHandler : HandlerBase
    {
        public InternalHandler() { Type = PurchaseType.Internal; }
        public override void Process(Request request) { request.Price *= 0.6; }
    }

    /// <summary>
    /// 只适于“邮购价格”的操作
    /// </summary>
    public class MailHandler : HandlerBase
    {
        public MailHandler() { Type = PurchaseType.Mail; }
        public override void Process(Request request) { request.Price *= 1.3; }
    }

    /// <summary>
    /// 只适于“折扣价”的操作
    /// </summary>
    public class DiscountHandler : HandlerBase
    {
        public DiscountHandler() { Type = PurchaseType.Discount; }
        public override void Process(Request request) { request.Price *= 0.9; }
    }

    /// <summary>
    /// 只适于“平价”的操作
    /// </summary>
    public class RegularHandler : HandlerBase
    {
        public RegularHandler() { Type = PurchaseType.Regular; }
    }

    #endregion

}
