using System;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.CoR.NonChain;
namespace MarvellousWorks.PracticalPattern.CoR.Tests.NonChain
{
    [TestClass]
    public class HandlerFixture
    {
        const double OriginalPrice = 100;
        Request mailRequest;

        [TestInitialize]
        public void Initialize()
        {
            mailRequest = new Request
            {
                Option = RequestOptions.Purchase,
                Price = 100,
                Type = PurchaseType.Mail
            };            
        }

        /// <summary>
        /// 验证购买过程处理CoR
        /// </summary>
        [TestMethod]
        public void TestPurchaseHandlerCoR()
        {
            Trace.Write("验证购买过程处理CoR");
            var head = new HandlerCoRFactory().CreateHandlerCoR(RequestOptions.Purchase);
            //  验证通过Lamada筛选出来的适用的Handler数量
            //  InternalHandler、MailPurchaseHandler、DiscountPurchaseHandler、RegularHandler 
            Assert.AreEqual<int>(4, head.Count());
            head.ToList().ForEach(x => x.HandleRequest(mailRequest));
            //  验证购买时邮购价格是原价的1.3倍
            Assert.AreEqual<double>(OriginalPrice * 1.3, mailRequest.Price);
        }

        /// <summary>
        /// 验证退货过程处理CoR
        /// </summary>
        [TestMethod]
        public void TestReturnHandlerCoR()
        {
            Trace.Write("验证退货过程处理CoR");
            var head = new HandlerCoRFactory().CreateHandlerCoR(RequestOptions.Return);
            //  验证通过Lamada筛选出来的适用的Handler数量
            //  InternalHandler、MailReturnHandler、DiscountReturnHandler、RegularHandler 
            Assert.AreEqual<int>(4, head.Count());
            head.ToList().ForEach(x => x.HandleRequest(mailRequest));
            //  验证邮购价格退货时只退还原价
            Assert.AreEqual<double>(OriginalPrice, mailRequest.Price);
        }

        /// <summary>
        /// 验证不能退货的情况
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void TestCannotReturnHandlerCoR()
        {
            Trace.Write("验证不能退货的情况CoR");
            var discountRequest = new Request()
            {
                Option = RequestOptions.Return,
                Price = OriginalPrice,
                Type = PurchaseType.Discount
            };
            new HandlerCoRFactory().CreateHandlerCoR(RequestOptions.Return).ToList().
                ForEach(x=>x.HandleRequest(discountRequest));
        }
    }
}
