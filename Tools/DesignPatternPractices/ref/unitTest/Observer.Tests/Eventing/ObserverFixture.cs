using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Observer.Eventing;
namespace MarvellousWorks.PracticalPattern.Observer.Tests.Eventing
{
    [TestClass]
    public class ObserverFixture
    {
        const string NewName = "joe";
        const string AnotherNewName = "peter";
        class A
        {
            public string Token { get; set; }
            public void OnNameChanged(object sender, UserEventArgs args)
            {
                Token = args.Name;
            }
        }

        class B
        {
            public string NewName { get; set; }
            public void OnNewNameNotified(object sender, UserEventArgs args)
            {
                NewName = args.Name;
            }
        }

        [TestMethod]
        public void Test()
        {
            var user = new User();
            var a1 = new A();
            var a2 = new A();
            var bList = new List<B> {new B(), new B(), new B()};
            user.NameChanged += a1.OnNameChanged;
            user.NameChanged += a2.OnNameChanged;
            bList.ForEach(x => user.NameChanged += x.OnNewNameNotified);

            // 确认未更新前的状态
            Assert.IsTrue(string.IsNullOrEmpty(a1.Token));
            Assert.IsTrue(string.IsNullOrEmpty(a2.Token));
            bList.ForEach(x=>Assert.IsTrue(string.IsNullOrEmpty(x.NewName)));

            // 更新
            user.Name = NewName;

            // 确认更新后的状态
            Assert.AreEqual<string>(NewName, a1.Token);
            Assert.AreEqual<string>(NewName, a2.Token);
            bList.ForEach(x => Assert.AreEqual<string>(NewName, x.NewName));

            // 调整观察者后，注销的观察者不会收到通知
            user.NameChanged -= a1.OnNameChanged;
            bList.ForEach(x => user.NameChanged -= x.OnNewNameNotified);
            user.Name = AnotherNewName;
            Assert.AreEqual<string>(NewName, a1.Token);
            Assert.AreEqual<string>(AnotherNewName, a2.Token);
            bList.ForEach(x => Assert.AreEqual<string>(NewName, x.NewName));
        }
    }
}
