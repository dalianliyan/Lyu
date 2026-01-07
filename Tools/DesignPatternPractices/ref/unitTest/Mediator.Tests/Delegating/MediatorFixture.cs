using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Mediator.Delegating;
namespace MarvellousWorks.PracticalPattern.Mediator.Tests.Delegating
{
    [TestClass]
    public class MediatorFixture
    {
        public class A : ColleagueBase<int>
        {
            public event EventHandler<DataEventArgs<int>> Changed;
            public override int Data
            {
                get{return base.Data;}
                set
                {
                    base.Data = value;
                    /// 把消息抛给作为中介的.NET事件机制
                    Changed(this, new DataEventArgs<int>(value));
                }
            }
        }
        public class B : ColleagueBase<int> { }
        public class C : ColleagueBase<int> { }

        [TestMethod]
        public void Test()
        {
            // 验证.NET事件对协作对象间的通知作用
            // 其中.NET事件机制作为隐含的Mediator对象出现
            var a = new A();
            var b = new B();
            var c = new C();
            a.Changed += b.OnChanged;
            a.Changed += c.OnChanged;
            a.Data = 20;
            Assert.AreEqual<int>(20, b.Data);
            Assert.AreEqual<int>(20, c.Data);

            // 更改协作关系
            a.Changed -= c.OnChanged;
            a.Data = 30;
            Assert.AreEqual<int>(30, b.Data);
            // C 因为不在新的协作关系之内，所以不变化
            Assert.AreEqual<int>(20, c.Data);

        }
    }
}
