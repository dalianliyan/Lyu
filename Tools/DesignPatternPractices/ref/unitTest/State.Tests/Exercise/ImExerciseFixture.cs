using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.State.Tests.Properties;
namespace MarvellousWorks.PracticalPattern.State.Tests.Exercise
{
    enum StateOption
    {
        Online,
        Offline,
        AppearOffline,
        Busy
    }

    class Wink { }      //  动漫表情
    class Message { }   //  消息

    #region 抽象State定义

    /// <summary>
    /// 模拟发送消息
    /// </summary>
    /// <remarks>
    ///     与一般状态模式示例不同的是，该示例涉及多个Context对象间的交互过程，而且状态较多、状态间协调关系复杂。<br/>
    ///     为此，实现上做了如下处理：
    ///     <list type="number">
    ///         <li>引入了创建型模式隔离具体State对象间的依赖关系</li>
    ///         <li>在IState定义上引入了决定是否能够交互的属性方法</li>
    ///     </list>
    /// </remarks>    
    interface IState
    {
        StateOption Option { get; }
        Account Account { get; }
        void Login();
        void Logoff();
        /// <summary>
        /// 只适于Online、Busy、AppearOffline三个状态间的转换, 而且不允许转换为现在所在的状态
        /// </summary>
        /// <param name="option">目标状态</param>
        void ChangeState(StateOption option);
        void SendMessage(Account account, Message message);
        void SendWink(Account account, Wink wink);
        bool CanReceiveMessage { get; }
        bool CanReceiveWink { get; }
    }

    abstract class StateBase : IState
    {
        public Account Account { get; protected set; }
        public StateOption Option { get; protected set; }

        protected StateBase(Account account, StateOption option)
        {
            if (account == null) throw new ArgumentNullException("account");
            Account = account;
            Option = option;
        }

        public virtual void Login()
        {
            this.Account.State = new StateFactory().Create(Account, StateOption.Online);
            TraceLogin(Account);
        }

        public virtual void Logoff()
        {
            this.Account.State = new StateFactory().Create(Account, StateOption.Offline);
            TraceLogoff(this.Account);
        }

        public virtual void SendMessage(Account account, Message message)
        {
            if (account == null) throw new ArgumentNullException("account");
            if (message == null) throw new ArgumentNullException("message");
            if (account.State.CanReceiveMessage)
                TraceSendMessage(this.Account, account);
            else
                throw new NotSupportedException();
        }

        public virtual void SendWink(Account account, Wink wink)
        {
            if (account == null) throw new ArgumentNullException("account");
            if (wink == null) throw new ArgumentNullException("wink");
            if (account.State.CanReceiveWink)
                TraceSendWink(this.Account, account);
            else
                throw new NotSupportedException();
        }
        public virtual void ChangeState(StateOption option)
        {
            if ((option == StateOption.Offline) || (option == Option))
                throw new NotSupportedException();
            Account.State = new StateFactory().Create(Account, option);
            TraceChangeState(Account, StateOption.Online, option);
        }

        public bool CanReceiveMessage { get; protected set; }
        public bool CanReceiveWink { get; protected set; }

        #region Trace Out

        static void TraceMessage(string formatter, params object[] parameters)
        {
            Trace.WriteLine(string.Format(formatter, parameters));
        }

        static void TraceLogin(Account account) { TraceMessage(Resource.LoginText, account.Name); }
        static void TraceLogoff(Account account) { TraceMessage(Resource.LogoffText, account.Name); }
        static void TraceChangeState(Account account, StateOption original, StateOption current) { TraceMessage(Resource.ChangeStateText, account.Name, original, current); }
        static void TraceSendMessage(Account from, Account to) { TraceMessage(Resource.SendMessageText, from.Name, to.Name); }
        static void TraceSendWink(Account from, Account to) { TraceMessage(Resource.SendWinkText, from.Name, to.Name); }

        #endregion
    }


    /// <summary>
    /// 构造State的工厂类型
    /// 隔离具体State对象间的依赖关系
    /// </summary>
    class StateFactory
    {
        public IState Create(Account account, StateOption option)
        {
            if (account == null) throw new ArgumentNullException("account");
            switch (option)
            {
                case StateOption.Online:
                    return new OnlineState(account);
                case StateOption.AppearOffline:
                    return new AppearOfflineState(account);
                case StateOption.Busy:
                    return new BusyState(account);
                case StateOption.Offline:
                    return new OfflineState(account);
                default:
                    throw new NotSupportedException();
            }
        }
    }

    #endregion

    #region Concrete State

    class OnlineState : StateBase
    {
        public OnlineState(Account account)
            : base(account, StateOption.Online)
        {
            CanReceiveMessage = true;
            CanReceiveWink = true;
        }

        public override void Login() { throw new NotSupportedException(); }
    }

    class OfflineState : StateBase
    {
        public OfflineState(Account account)
            : base(account, StateOption.Online)
        {
            CanReceiveMessage = false;
            CanReceiveWink = false;
        }

        public override void Logoff() { throw new NotSupportedException(); }
        public override void SendMessage(Account account, Message message) { throw new NotSupportedException(); }
        public override void SendWink(Account account, Wink wink) { throw new NotSupportedException(); }
        public override void ChangeState(StateOption option) { throw new NotSupportedException(); }
    }

    class AppearOfflineState : StateBase
    {
        public AppearOfflineState(Account account)
            : base(account, StateOption.AppearOffline)
        {
            CanReceiveMessage = false;
            CanReceiveWink = false;
        }

        public override void Login() { throw new NotSupportedException(); }
        public override void SendWink(Account account, Wink wink) { throw new NotSupportedException(); }
    }

    class BusyState : StateBase
    {
        public BusyState(Account account)
            : base(account, StateOption.Busy)
        {
            CanReceiveMessage = true;
            CanReceiveWink = false;
        }

        public override void Login() { throw new NotSupportedException(); }
    }

    #endregion

    /// <summary>
    /// Context对象
    /// </summary>
    class Account
    {
        public string Name { get; private set; }
        public IState State { get; set; }

        public Account(string name)
        {
            Name = name;
            // 默认状态为离线状态
            State = new OfflineState(this);
        }

        #region 应用State模式完成各种与状态有关的操作

        public void Login() { State.Login(); }
        public void Logoff() { State.Logoff(); }
        public void SendMessage(Account account, Message message) { State.SendMessage(account, message); }
        public void SendWink(Account account, Wink wink) { State.SendWink(account, wink); }
        public void ChangeState(StateOption option) { State.ChangeState(option); }

        #endregion
    }


    [TestClass]
    public class ImExerciseFixture
    {
        Account accountA;
        Account accountB;
        Account accountC;

        [TestInitialize]
        public void Initialize()
        {
            accountA = new Account("A");
            accountB = new Account("B");
            accountC = new Account("C");
        }

        [TestMethod]
        public void TestInitAccount()
        {
            Assert.IsNotNull(accountA);
            Assert.IsNotNull(accountB);
            Assert.IsNotNull(accountC);

            //  验证初始状态的正确性
            Assert.IsInstanceOfType(accountA.State, typeof (OfflineState));
            Assert.IsInstanceOfType(accountB.State, typeof(OfflineState));
            Assert.IsInstanceOfType(accountC.State, typeof(OfflineState));
        }

        [TestMethod]
        public void TestOfflineState()
        {
            //  offline的用户不能继续logoff
            try
            {
                accountA.Logoff();
            }
            catch(NotSupportedException)
            {
            }

            accountA.Login();
            Assert.IsInstanceOfType(accountA.State, typeof(OnlineState));

            #region  Offline状态下用户不可以发消息和动漫表情，不能接受消息和动漫表情

            Assert.IsFalse(accountB.State.CanReceiveMessage);
            Assert.IsFalse(accountB.State.CanReceiveWink);

            try
            {
                accountB.SendMessage(accountC, new Message());
            }
            catch(NotSupportedException)
            {
            }

            try
            {
                accountB.SendWink(accountC, new Wink());
            }
            catch (NotSupportedException)
            {
            }

            try
            {
                accountB.ChangeState(StateOption.AppearOffline);
            }
            catch (NotSupportedException)
            {
            }

            #endregion
        }

        [TestMethod]
        public void TestOnlineState()
        {
            accountA.Login();
            accountB.Login();
            Assert.IsInstanceOfType(accountA.State, typeof(OnlineState));
            Assert.IsInstanceOfType(accountB.State, typeof(OnlineState));
            
            #region Online状态下用户可以发消息和动漫表情，可以接受消息和动漫表情

            Assert.IsTrue(accountA.State.CanReceiveMessage);
            Assert.IsTrue(accountA.State.CanReceiveWink);

            accountA.SendMessage(accountB, new Message());
            accountA.SendWink(accountB, new Wink());

            #endregion

            accountB.Logoff();
            Assert.IsInstanceOfType(accountB.State, typeof(OfflineState));

        }

        [TestMethod]
        public void TestBusyState()
        {
            accountA.Login();
            accountB.Login();

            accountB.ChangeState(StateOption.Busy);
            Assert.IsInstanceOfType(accountB.State, typeof(BusyState));

            accountB.ChangeState(StateOption.AppearOffline);
            Assert.IsInstanceOfType(accountB.State, typeof(AppearOfflineState));

            accountB.ChangeState(StateOption.Online);
            Assert.IsInstanceOfType(accountB.State, typeof(OnlineState));

            accountB.ChangeState(StateOption.Busy);

            #region Busy状态的用户可以发送消息和动漫表情，可以接受消息但不能接受动漫表情

            Assert.IsTrue(accountB.State.CanReceiveMessage);
            Assert.IsFalse(accountB.State.CanReceiveWink);

            accountA.SendMessage(accountB, new Message());
            accountB.SendMessage(accountA, new Message());
            accountB.SendWink(accountA, new Wink());
            try
            {
                accountA.SendWink(accountB, new Wink());
            }
            catch(NotSupportedException)
            {
            }

            #endregion

            accountB.Logoff();
            Assert.IsInstanceOfType(accountB.State, typeof(OfflineState));

            #region 对于Offline的用户只能通过登录（Login）进入Online状态，然后再选择进入Busy或AppearOffline状态

            try
            {
                accountC.ChangeState(StateOption.Busy);   
            }
            catch(NotSupportedException)
            {
            }

            #endregion

        }

        [TestMethod]
        public void TestAppearOffline()
        {
            accountA.Login();
            accountB.Login();

            accountB.ChangeState(StateOption.AppearOffline);
            Assert.IsInstanceOfType(accountB.State, typeof(AppearOfflineState));

            accountB.ChangeState(StateOption.Busy);
            Assert.IsInstanceOfType(accountB.State, typeof(BusyState));

            accountB.ChangeState(StateOption.Online);
            Assert.IsInstanceOfType(accountB.State, typeof(OnlineState));

            accountB.ChangeState(StateOption.AppearOffline);

            #region 状态下能够发送消息但不能发送动漫表情，不可以接受消息和动漫表情

            Assert.IsFalse(accountB.State.CanReceiveMessage);
            Assert.IsFalse(accountB.State.CanReceiveWink);

            accountB.SendMessage(accountA, new Message());
            try
            {
                accountA.SendWink(accountB, new Wink());
            }
            catch (NotSupportedException)
            {
            }

            try
            {
                accountA.SendMessage(accountB, new Message());
            }
            catch (NotSupportedException)
            {
            }

            try
            {
                accountB.SendWink(accountA, new Wink());
            }
            catch (NotSupportedException)
            {
            }

            #endregion

            accountB.Logoff();
            Assert.IsInstanceOfType(accountB.State, typeof (OfflineState));

            #region 对于Offline的用户只能通过登录（Login）进入Online状态，然后再选择进入Busy或AppearOffline状态

            try
            {
                accountC.ChangeState(StateOption.AppearOffline);
            }
            catch (NotSupportedException)
            {
            }

            #endregion

        }
    }
}
