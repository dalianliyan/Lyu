using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Concept.DependencyInjection;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.DependencyInjection.Attributer
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple=true)]
    sealed class DecoratorAttribute : Attribute
    {
        /// <summary>
        /// 实现客户类型实际需要的抽象类型的实体类型实例，即待注入到客户类型的内容
        /// </summary>
        public readonly object Injector;
        readonly Type type;

        public DecoratorAttribute(Type type) 
        {
            if (type == null) throw new ArgumentNullException("type");
            this.type = type;
            Injector = (new Assembler()).Create(this.type);
        }

        /// <summary>
        /// 客户类型需要的抽象对象类型
        /// </summary>
        public Type Type { get { return this.type; } }
    }

    /// <summary>
    /// 用户帮助客户类型和客户程序获取其Attribute定义中需要的抽象类型实例
    /// </summary>
    static class AttributeHelper
    {
        public static T Injector<T>(object target)
            where T : class
        {
            if (target == null) throw new ArgumentNullException("target");
            return (T)(((DecoratorAttribute[])
                target.GetType().GetCustomAttributes(typeof(DecoratorAttribute), false))
                .Where(x => x.Type == typeof(T))
                .FirstOrDefault()
                .Injector);
        }
    }

    [Decorator(typeof(ITimeProvider))]
    class Client
    {
        public int GetYear()
        {
            // 与其他方式注入不同的是，这里使用的ITimeProvider来自自己的Attribute
            var provider = AttributeHelper.Injector<ITimeProvider>(this);
            return provider.CurrentDate.Year;
        }
    }

    [TestClass]
    public class TestClient
    {
        [TestMethod]
        public void Test()
        {
            Assert.IsTrue(new Client().GetYear() > 0);
        }
    }
}
