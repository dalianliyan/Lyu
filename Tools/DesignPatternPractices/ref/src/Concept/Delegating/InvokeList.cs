using System;
using System.Linq;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Concept.Delegating
{
    public class InvokeList
    {
        List<StringAssignmentEventHandler> handlers;
        string[] message = new string[3];

        #region Common Method Constructor
        public InvokeList()
        {
            // 绑定一组抽象方法
            handlers = new List<StringAssignmentEventHandler>()
                           {
                               AppendHello,
                               AppendComma,
                               AppendWorld
                           };
        }

        // 具体操作方法
        public void AppendHello() { message[0] = "hello"; }
        public void AppendComma() { message[1] = ","; }
        public void AppendWorld() { message[2] = "world"; }

        public void Invoke()
        {
            handlers.ForEach(x=>x());
        }
        #endregion

        #region Anonymous Method Constructor Common Version
        //public InvokeList()
        //{
        //    // 绑定一组抽象方法
        //    handlers = new List<StringAssignmentEventHandler>();
        //    StringAssignmentEventHandler h1 = delegate { message[0] = "hello"; };
        //    StringAssignmentEventHandler h2 = delegate { message[1] = ","; };
        //    StringAssignmentEventHandler h3 = delegate { message[2] = "world"; };
        //    handlers.Add(h1);
        //    handlers.Add(h2);
        //    handlers.Add(h3);
        //}

        //public void Invoke()
        //{
        //    foreach (StringAssignmentEventHandler handler in handlers)
        //        handler();
        //}
        #endregion

        #region Anonymous Method Constructor Simple Version
        
        //public MulticastDelegateInvoker()  
        //{
        //    StringAssignmentEventHandler handler = null;
        //    handler += new StringAssignmentEventHandler(AppendHello);
        //    handler += new StringAssignmentEventHandler(AppendComma);
        //    handler += new StringAssignmentEventHandler(AppendWorld);
        //    handler.Invoke();
        //}

        #endregion

        public string this[int index] { get { return message[index]; } }
    }
}
