using System;
using System.Collections.Generic;
using System.Configuration;
using System.Collections.Specialized;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.FactoryMethod;
namespace MarvellousWorks.PracticalPattern.FactoryMethod.Tests.Classic
{
    public class Assembler
    {
        /// <summary>
        /// 配置节名称
        /// </summary>
        const string SectionName = "factoryMethod.customFactories";

        /// <summary>
        /// 保存“抽象类型/实体类型”对应关系的字典
        /// </summary>
        static Dictionary<Type, Type> dictionary = new Dictionary<Type, Type>();

        static Assembler()
        {
            // 通过配置文件加载相关“抽象产品类型”/ “实体产品类型”映射关系
            NameValueCollection collection =
                (NameValueCollection)ConfigurationManager.GetSection(SectionName);
            for (int i = 0; i < collection.Count; i++)
            {
                string target = collection.GetKey(i);
                string source = collection[i];
                dictionary.Add(Type.GetType(target), Type.GetType(source));
            }
        }

        /// <summary>
        /// 根据客户程序需要的抽象类型选择相应的实体类型，并返回类型实例
        /// </summary>
        /// <typeparam name="T">抽象类型（抽象类/接口/或者某种基类）</typeparam>
        /// <returns>实体类型实例</returns>
        public object Create(Type type)     // 主要用于非泛型方式调用
        {
            if ((type == null) || !dictionary.ContainsKey(type)) throw new NullReferenceException();
            return Activator.CreateInstance(dictionary[type]);
        }

        /// <typeparam name="T">抽象类型（抽象类/接口/或者某种基类）</typeparam>
        public T Create<T>()    // 主要用于泛型方式调用
        {
            return (T)Create(typeof(T));
        }

        public void Assembly(Client client)
        {
            client.Factory = Create<IFactory>();
        }
    }

    /// <summary>
    /// 抽象的工厂类型特性描述
    /// </summary>
    public interface IFactory
    {
        IProduct Create();  //  每个工厂均具有的功能——构造产品
    }

    /// <summary>
    /// 具体工厂类型
    /// </summary>
    public class FactoryA : IFactory
    {
        public IProduct Create()
        {
            return new ProductA();
        }
    }

    /// <summary>
    /// 具体工厂类型
    /// </summary>
    public class FactoryB : IFactory
    {
        public IProduct Create()
        {
            return new ProductB();
        }
    }


    [TestClass]
    public class Client
    {
        public IFactory Factory { get; set; }   // setter injection
        public IProduct Product { get { return Factory.Create(); } }

        [TestMethod]
        public void FactoryWithAssembler()
        {
            var client = new Client();
            new Assembler().Assembly(client);
            Assert.IsInstanceOfType(client.Factory, typeof(FactoryA));
            Assert.IsInstanceOfType(client.Product, typeof(ProductA));
        }
    }

}
