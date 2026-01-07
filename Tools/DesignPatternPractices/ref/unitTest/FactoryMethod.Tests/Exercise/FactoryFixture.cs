using System;
using F = MarvellousWorks.PracticalPattern.FactoryMethod;
using Microsoft.VisualStudio.TestTools.UnitTesting; 
namespace MarvellousWorks.PracticalPattern.FactoryMethod.Tests.Exercise
{
 
    [TestClass]
    public class FactoryFixture
    {
        interface IFactory : F.IFactory
        {
            #region factory method
            TTarget Create<TTarget>(object[] parametes);
            TTarget Create<TTarget>(string name, object[] parametes);
            #endregion
        }
        
        class Factory : F.Factory, IFactory 
        {
            public TTarget Create<TTarget>(object[] parametes)
            {
                return (TTarget)Activator.CreateInstance(registry[typeof(TTarget)], parametes);
            }

            public TTarget Create<TTarget>(string name, object[] parametes)
            {
                return (TTarget)Activator.CreateInstance(registry[typeof(TTarget), name], parametes);
            }
        }

        [TestMethod]
        public void CreateInstance()
        {
            var factory = new Factory()
                .RegisterType<IFruit, Apple>()
                .RegisterType<IFruit, Orange>("o")
                .RegisterType<IVehicle, Bicycle>()
                .RegisterType<IVehicle, Bicycle>("a")
                .RegisterType<IVehicle, Train>("b")
                .RegisterType<IVehicle, Car>("c")
                .RegisterType<IEntry, EntryWithName>()
                .RegisterType<IEntry, EntryWithName>("n")
                .RegisterType<IEntry, EntryWithNameAndAgeAndTitle>("nat");

            #region 构造函数无参数的类型

            Assert.IsInstanceOfType(factory.Create<IFruit>(), typeof(Apple));
            Assert.IsInstanceOfType(factory.Create<IFruit>("o"), typeof(Orange));

            Assert.IsInstanceOfType(factory.Create<IVehicle>(), typeof(Bicycle));
            Assert.IsInstanceOfType(factory.Create<IVehicle>("a"), typeof(Bicycle));
            Assert.IsInstanceOfType(factory.Create<IVehicle>("b"), typeof(Train));
            Assert.IsInstanceOfType(factory.Create<IVehicle>("c"), typeof(Car));

            #endregion

            #region 构造函数带参数的类型

            //  转换为新的扩展接口形式
            var f = (IFactory) factory;

            //  使用扩展的功能

            var e1 = f.Create<IEntry>(new object[] { "joe" });
            Assert.IsInstanceOfType(e1, typeof(EntryWithName));
            Assert.AreEqual<string>("joe", ((EntryWithName) e1).Name);

            var e2 = f.Create<IEntry>("nat", new object[] { "joe", 20});
            Assert.IsInstanceOfType(e2, typeof(EntryWithNameAndAgeAndTitle));
            Assert.AreEqual<string>("joe", ((EntryWithName)e2).Name);
            Assert.AreEqual<int>(20, ((EntryWithNameAndAgeAndTitle)e2).Age);
            Assert.AreEqual<string>(EntryWithNameAndAgeAndTitle.DefaultTitle, ((EntryWithNameAndAgeAndTitle)e2).Title);

            #endregion
        }
    }
}
