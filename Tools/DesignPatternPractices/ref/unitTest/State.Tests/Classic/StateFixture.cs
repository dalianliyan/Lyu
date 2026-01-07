using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.State.Classic;
namespace MarvellousWorks.PracticalPattern.State.Tests.Classic
{
    [TestClass]
    public class StateFixture
    {
        /// <summary>
        /// 具体的State类型
        /// </summary>
        class OpenState : IState
        {
            public void Open() { throw new NotSupportedException(); }
            public void Close() { } // pass
            public void Query() { } // pass
        }

        class ClosedState : IState
        {
            public void Open() { }
            public void Close(){}
            public void Query() { throw new NotSupportedException(); }
        }

        /// <summary>
        /// 具体Context类型
        /// </summary>
        class Connection : ContextBase { }

        Connection connection;
        [TestInitialize]
        public void Initialize() { connection = new Connection(); }

        /// <summary>
        /// 根据状态判定为不合规的流程
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void TryOnAnOpenConnection()
        {
            connection.State = new OpenState();
            connection.Open();
        }

        /// <summary>
        /// 根据状态判定为不合规的流程
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void TryOnAnClosedConnection()
        {
            connection.State = new ClosedState();
            connection.Query();
        }

        /// <summary>
        /// 根据状态判定为合规的流程
        /// </summary>
        [TestMethod]
        public void TrySupportOperations()
        {
            connection.State = new OpenState();
            connection.Query();
            connection.Close();

            connection.State = new ClosedState();
            connection.Close();
            connection.Open();
        }
    }
}
