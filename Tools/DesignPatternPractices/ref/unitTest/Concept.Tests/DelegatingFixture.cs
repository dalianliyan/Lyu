using System.Diagnostics;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Concept.Delegating;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    [TestClass]
    public class DelegatingFixture
    {
        [TestMethod]
        public void SimpleAysncInvoke()
        {
            new AsyncInvoker();
            Thread.Sleep(3000);
        }

        [TestMethod]
        public void AysncInvokeList()
        {
            var list = new InvokeList();
            list.Invoke();
            Assert.AreEqual<string>("hello,world", list[0] + list[1] + list[2]);
        }

        [TestMethod]
        public void MulticastDelegateInvoke()
        {
            var invoker = new MulticastDelegateInvoker();
            Assert.AreEqual<string>("hello,world", invoker[0] + invoker[1] + invoker[2]);
        }

        [TestMethod]
        public void EventMonitorSimulate()
        {
            var order1 = new Order();
            order1.Create();        // add      1
            order1.ChangeDate();    // modify   1
            order1.ChangeDate();    // modify   2
            order1.ChangeOwner();   // modify   3

            var order2 = new Order();
            order2.Create();        //  add     2
            order2.ChangeOwner();   //  modify  4
            order2.ChangeId();      //  modify  still 4

            Assert.AreEqual<int>(2, EventMonitor.AddedTimes);
            Assert.AreEqual<int>(4, EventMonitor.ModifiedTimes);
        }
    }



    public delegate int TwoParametersEventHandler(int x, int y);
    public delegate int DualParametersEventHandler(int pX, int pY);
    public delegate int AddEventHandler(int a, int b);

    public class Apple{}
    public class User{}
    public delegate int JustTwoParametersEventHandler(int pX, int pY);
    public delegate Apple ChooseBetterAppleEventHandler(Apple a1, Apple a2);
    public delegate User GetElderUserEventHandler(User user1, User user2);

    [TestClass]
    public class LamadaTimerFixture
    {
        class AsyncInvoker
        {
            /// <summary>
            /// Lamada方式定义Timer所需的回调委托过程
            /// </summary>
            /// <remarks>
            /// 这里(x) => Trace.WriteLine(x as string)就是用Lamada语法定义的委托回调内容
            /// </remarks>
            public AsyncInvoker()
            {
                Trace.WriteLine("method");
                new Timer((x) => Trace.WriteLine(x), "slow", 2500, 2500);
                new Timer((x) => Trace.WriteLine(x), "fast", 2000, 2000);
            }
        }

        [TestMethod]
        public void SimpleAysncInvoke()
        {
            new AsyncInvoker();
            Thread.Sleep(3000);
        }
    }

}
