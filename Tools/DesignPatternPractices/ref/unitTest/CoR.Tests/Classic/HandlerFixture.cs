using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.CoR.Classic;
namespace MarvellousWorks.PracticalPattern.CoR.Tests.Classic
{
    [TestClass]
    public class HandlerFixture
    {
        [TestMethod]
        public void Test()
        {
            Trace.WriteLine("Process via whole CoR");
            // 生成并组装链式的结构  internal-> discount-> mail-> regular-> null
            var head = new InternalHandler()
            {
                Successor = new DiscountHandler()
                {
                    Successor = new MailHandler()
                    {
                        Successor = new RegularHandler()
                    }                
                }
            };

            var request = new Request() { Price = 20, Type = PurchaseType.Mail };
            head.HandleRequest(request);
            Assert.AreEqual<double>(20 * 1.3, request.Price);
            request = new Request() { Price = 20, Type = PurchaseType.Discount };
            head.HandleRequest(request);
            Assert.AreEqual<double>(20 * 0.9, request.Price);

            Trace.WriteLine("\n\nProcess via rearranged CoR");
            //  重新组织链表结构， 新链表为internal-> mail-> regular-> null
            //  此时，head指向internal
            var discountHandler = head.Successor;
            head.Successor = head.Successor.Successor;  // 短路掉Discount Handler
            discountHandler = null;
            request = new Request() { Price = 20, Type = PurchaseType.Discount };
            head.HandleRequest(request);

            // 价格没有折扣打折，确认被短路的部分无法生效
            //  验证CoR可以通过动态维护链表调整处理对象（Handler）的组织结构
            Assert.AreEqual<double>(20, request.Price);    
        }
    }
}
