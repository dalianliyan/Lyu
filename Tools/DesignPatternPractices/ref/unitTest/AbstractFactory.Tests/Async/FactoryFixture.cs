using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.AbstractFactory.Async;
namespace MarvellousWorks.PracticalPattern.AbstractFactory.Tests.Async
{
    [TestClass]
    public class FactoryFixture
    {
        class ConcreteProduct : IProduct{}
        class ConcreteFactory : IFactoryWithNotifier
        {
            /// <summary>
            /// 同步构造过程
            /// </summary>
            /// <returns></returns>
            public IProduct Create() { return new ConcreteProduct(); }

            /// <summary>
            /// 异步构造过程
            /// </summary>
            /// <param name="callback"></param>
            public void Create(Action<IProduct> callback)
            {
                callback(Create());
            }
        }

        class Subscribe
        {
            public IProduct Product { get; set; }
        }

        [TestMethod]
        public void Test()
        {
            IFactoryWithNotifier factory = new ConcreteFactory();
            Subscribe subcriber = new Subscribe();
            Assert.IsNull(subcriber.Product);
            factory.Create((x) => { subcriber.Product = x; });
            Assert.IsNotNull(subcriber.Product);
            Assert.IsTrue(subcriber.Product is ConcreteProduct);
        }
    }
}
