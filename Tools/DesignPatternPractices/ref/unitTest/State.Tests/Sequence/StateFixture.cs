using System;
using MarvellousWorks.PracticalPattern.State.Sequence;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.State.Tests.Sequence
{
    [TestClass]
    public class StateFixture
    {
        /// <summary>
        /// 具体的State类型
        /// </summary>
        class OpenState : IState
        {
            public void Open(IContext context) { throw new NotSupportedException(); }
            public void Close(IContext context) { context.State = new CloseState(); }
            public void Query(IContext context) { } 
        }
        class CloseState : IState
        {
            public void Open(IContext context) { context.State = new OpenState(); }
            public void Close(IContext context) { }
            public void Query(IContext context) { throw new NotSupportedException(); }
        }

        /// <summary>
        /// 具体Context类型
        /// </summary>
        class Connection : ContextBase
        {
            public Connection() { State = new CloseState(); }
        }

        Connection connection;

        [TestInitialize]
        public void Initialize()
        {
            connection = new Connection();
        }

        [TestMethod]
        public void NormalStateSequence()
        {
            var random = new Random();
            for(var i=0; i<100; i++)
            {
                connection.Open();
                for (var j = 0; j < random.Next() % 23; j++)
                    connection.Query();
                connection.Close();
            }
        }

        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void AbnormalStateSequence()
        {
            connection.Close();
            connection.Query();
        }

    }

}
