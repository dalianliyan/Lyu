using System;

namespace MarvellousWorks.PracticalPattern.CoR.NonChain
{
    #region 具体操作类型

    /// <summary>
    /// 适于“内部价格”的购买和退货操作
    /// </summary>
    public class InternalHandler : HandlerBase
    {
        public InternalHandler() 
        {
            Type = PurchaseType.Internal;
            Option = RequestOptions.Purchase | RequestOptions.Return;
        }
        public override void Process(Request request) { request.Price *= 0.6; }
    }

    /// <summary>
    /// 适于“邮购价格”购买操作
    /// </summary>
    public class MailPurchaseHandler : HandlerBase
    {
        public MailPurchaseHandler() 
        { 
            Type = PurchaseType.Mail;
            Option = RequestOptions.Purchase; 
        }
        public override void Process(Request request) { request.Price *= 1.3; }
    }

    /// <summary>
    /// 适于“邮购价格”退货操作
    /// </summary>
    public class MailReturnHandler : HandlerBase
    {
        public MailReturnHandler()
        {
            Type = PurchaseType.Mail;
            Option = RequestOptions.Return;
        }
    }


    /// <summary>
    /// 适于“折扣价”购买操作
    /// </summary>
    public class DiscountPurchaseHandler : HandlerBase
    {
        public DiscountPurchaseHandler() 
        { 
            Type = PurchaseType.Discount;
            Option = RequestOptions.Purchase; 
        }
        public override void Process(Request request) { request.Price *= 0.9; }
    }

    /// <summary>
    /// 适于“折扣价”退货操作
    /// </summary>
    public class DiscountReturnHandler : HandlerBase
    {
        public DiscountReturnHandler()
        {
            Type = PurchaseType.Discount;
            Option = RequestOptions.Return;
        }
        /// <summary>
        /// 折扣价促销商品不退不换
        /// </summary>
        /// <param name="request"></param>
        public override void Process(Request request){throw new NotSupportedException();}
    }

    /// <summary>
    /// 适于“平价”购买和退货操作
    /// </summary>
    public class RegularHandler : HandlerBase
    {
        public RegularHandler()
        {
            Type = PurchaseType.Regular;
            Option = RequestOptions.Purchase | RequestOptions.Return;
        }
    }

    #endregion

}
