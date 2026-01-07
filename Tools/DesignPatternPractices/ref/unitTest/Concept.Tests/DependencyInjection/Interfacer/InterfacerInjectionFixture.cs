using System;
using MarvellousWorks.PracticalPattern.Concept.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.DependencyInjection.Interfacer
{
    [TestClass]
    public class InterfacerInjectionFixture
    {
        /// <summary>
        /// 通过泛型参数实现接口注入
        /// </summary>
        /// <typeparam name="T">待注入的接口类型</typeparam>
        class Client<T> : ITimeProvider
            where T : ITimeProvider
        {
            /// <summary>
            /// 与设值方式相似的注入入口
            /// </summary>
            public T Provider { get; set; }

            /// <summary>
            /// 类似传统接口注入的实现方式
            /// </summary>
            public DateTime CurrentDate
            {
                get { return Provider.CurrentDate; }
            }
        }

        [TestMethod]
        public void Test()
        {
            var client = new Client<ITimeProvider>()
                             {
                                 Provider = (new Assembler().Create<ITimeProvider>())
                             };

            //  验证设值方式注入的内容
            Assert.IsNotNull(client.Provider);
            Assert.IsInstanceOfType(client.Provider, typeof(SystemTimeProvider));

            //  验证注入的接口是否可用
            Assert.IsInstanceOfType(client.Provider.CurrentDate, typeof(DateTime));

            //  验证是否满足传统接口注入的要求
            Assert.IsTrue(typeof(ITimeProvider).IsAssignableFrom(client.GetType()));
        }
    }
}
