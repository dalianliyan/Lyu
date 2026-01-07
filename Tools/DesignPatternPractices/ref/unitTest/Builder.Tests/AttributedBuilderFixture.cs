using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Builder.Attributed;
namespace MarvellousWorks.PracticalPattern.Builder.Tests
{
    [TestClass]
    public class AttributedBuilderFixture
    {
        #region 采用BuildStep定义装配过程的测试类型

        /// <summary>
        /// 用于显示测试结果的委托
        /// </summary>
        /// <remarks>
        ///     Action<string>是实际处理Build Step操作内容的委托
        /// </remarks>
        static Action<string, Action<string>> buildPartHandler = (x, y) =>
        {
            Trace.WriteLine("add " + x);
            y(x);
        };

        class Car
        {
            public IList<string> Parts { get; private set; }
            public Car(){Parts = new List<string>();}

            /// <summary>
            /// 为汽车添加轮胎
            /// </summary>
            /// <remarks>
            ///     Attributed Builder第二个执行的Setp
            ///     执行4次，即为每辆汽车装配增加4个轮胎
            /// </remarks>
            [BuildStep(2, 4)]
            public void AddWheel() { buildPartHandler("wheel", Parts.Add); }

            /// <summary>
            /// 为汽车装配引擎
            /// </summary>
            /// <remarks>
            ///     没有通过Attribute标注的内容，因此不会被Attributed Builder执行
            /// </remarks>
            public void AddEngine() { buildPartHandler("engine", Parts.Add); }

            /// <summary>
            /// 为汽车装配车身
            /// </summary>
            /// <remarks>
            ///     Attributed Builder第一个执行的Setp
            ///     执行1次，即为每辆汽车装配增加1个车身
            /// </remarks>
            [BuildStep(1)]
            public void AddBody() { buildPartHandler("body", Parts.Add); }
        }

        class Computer
        {
            public string Bus { get; private set; }
            public string Monitor { get; private set; }
            public string Disk { get; private set; }
            public string Memory { get; private set; }
            public string Keyboard { get; private set; }
            public string Mouse { get; private set; }

            /// <summary>
            /// 缓存Computer类型的所有Property
            /// </summary>
            static PropertyInfo[] properties = typeof (Computer).GetProperties();

            /// <summary>
            /// 获得Computer类型指定名称Property的Setter方法委托
            /// </summary>
            /// <param name="target">Computer类型实例</param>
            /// <param name="name">Property名称</param>
            /// <returns>指定名称Property的Setter方法委托</returns>
            static Action<string> GetSetter(Computer target, string name)
            {
                var property = properties.Where(x => string.Equals(x.Name, name)).FirstOrDefault();
                return x => property.SetValue(target, x, null);
            }

            [BuildStep(1)]
            public void LayoutBus()
            {
                buildPartHandler("bus", GetSetter(this, "Bus"));
            }

            [BuildStep(2)]
            public void AddDiskAndMemory()
            {
                buildPartHandler("disk", GetSetter(this, "Disk"));
                buildPartHandler("16G memory", GetSetter(this, "Memory"));
            }

            [BuildStep(3)]
            public void AddUserInputDevice()
            {
                buildPartHandler("USB mouse", GetSetter(this, "Mouse"));
                buildPartHandler("keyboard", GetSetter(this, "Keyboard"));
            }

            [BuildStep(4)]
            public void ConnectMonitor()
            {
                buildPartHandler("monitor", GetSetter(this, "Monitor"));
            }
        }

        #endregion

        [TestMethod]
        public void BuildComputerByAttributeDirection()
        {
            Trace.WriteLine("\nassembly computer");
            var computer = new Computer();
            Assert.IsNull(computer.Keyboard);
            Assert.IsNull(computer.Memory);
            computer = new Builder<Computer>().BuildUp();
            Assert.IsNotNull(computer.Bus);
            Assert.IsNotNull(computer.Monitor);
            Assert.IsNotNull(computer.Disk);
            Assert.IsNotNull(computer.Memory);
            Assert.IsNotNull(computer.Keyboard);
            Assert.IsNotNull(computer.Mouse);
        }

        [TestMethod]
        public void BuildCarByAttributeDirection()
        {
            Trace.WriteLine("build car");
            var car = new Builder<Car>().BuildUp();
            Assert.IsNotNull(car);
            Assert.IsFalse(car.Parts.Contains("engine")); // 不会被执行的内容
            Assert.AreEqual<string>("body", car.Parts.ElementAt(0));
            for (var i = 1; i <= 4; i++)
                Assert.AreEqual<string>("wheel", car.Parts.ElementAt(i));
        }
    }
}

