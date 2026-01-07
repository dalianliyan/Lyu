using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Practices.Unity;
using Microsoft.Practices.Unity.InterceptionExtension;
namespace MarvellousWorks.PracticalPattern.Decorator.Tests.Interception
{
    [TestClass]
    public class MsilDecoratorInjectionFixture
    {
        /// <summary>
        /// 待装饰的目标类型接口
        /// </summary>
        public interface IBizObject
        {
            /// <summary>
            /// 待装饰的内容只要定义在接口即可，不用逐个定义在每个实体类上
            /// </summary>
            /// <returns></returns>
            [Log]
            [Security]
            int GetValue();

            int GetYear();
        }

        /// <summary>
        /// 待装饰的目标类型
        /// </summary>
        public class BizObject : IBizObject
        {
            public static readonly int Value = new Random().Next();

            /// <summary>
            /// 接口中已经被标注为需要装饰的方法
            /// </summary>
            /// <returns></returns>
            public int GetValue(){return Value;}

            /// <summary>
            /// 接口中声明为无需装饰的方法
            /// </summary>
            /// <returns></returns>
            public int GetYear()
            {
                return DateTime.Now.Year;    
            }
        }

        #region 自定义的装饰属性

        public class LogAttribute : HandlerAttribute, ICallHandler
        {
            public override ICallHandler CreateHandler(IUnityContainer container){return this;}
            public IMethodReturn Invoke(IMethodInvocation input, GetNextHandlerDelegate getNext)
            {
                Trace.WriteLine("log...");
                return getNext()(input, getNext);
            }
        }

        public class SecurityAttribute : HandlerAttribute, ICallHandler
        {
            public override ICallHandler CreateHandler(IUnityContainer container) { return this; }
            public IMethodReturn Invoke(IMethodInvocation input, GetNextHandlerDelegate getNext)
            {
                Trace.WriteLine("machine name");
                Trace.WriteLine("domain name");
                return getNext()(input, getNext);
            }
        }

        #endregion

        [TestMethod]
        public void TestInjectionType()
        {
            var container =
                new UnityContainer().AddNewExtension<Microsoft.Practices.Unity.InterceptionExtension.Interception>();

            //  Unity动态编译并增加装饰类型
            container.RegisterType<IBizObject, BizObject>().
                Configure<Microsoft.Practices.Unity.InterceptionExtension.Interception>().
                SetInterceptorFor<IBizObject>(new InterfaceInterceptor());

            var bizObject = container.Resolve<IBizObject>();

            //  调用装饰后的方法
            Trace.WriteLine("\nInvoke GetValue()\n-----------------");
            Assert.AreEqual<int>(BizObject.Value, bizObject.GetValue());

            //  调用未装饰的方法
            Trace.WriteLine("\nInvoke GetValue()\n-----------------");
            Assert.AreEqual<int>(DateTime.Now.Year, bizObject.GetYear());
        }
    }
}
