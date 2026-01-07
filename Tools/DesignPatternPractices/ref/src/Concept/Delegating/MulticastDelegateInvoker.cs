using System;
using System.Collections.Generic;

namespace MarvellousWorks.PracticalPattern.Concept.Delegating
{
    public class MulticastDelegateInvoker
    {
        string[] message = new string[3];

        //public MulticastDelegateInvoker()
        //{
        //    StringAssignmentEventHandler handler = null;
        //    handler += new StringAssignmentEventHandler(AppendHello);
        //    handler += new StringAssignmentEventHandler(AppendComma);
        //    handler += new StringAssignmentEventHandler(AppendWorld);
        //    handler.Invoke();
        //}


        #region Lamada

        public MulticastDelegateInvoker()
        {
            List<Action> handler = new List<Action>
                                       {
                                           () => { message[0] = "hello"; },
                                           () => { message[1] = ","; },
                                           () => { message[2] = "world"; }
                                       };
            handler.ForEach(x => x());
        }
        #endregion

        public void AppendHello() { message[0] = "hello"; }
        public void AppendComma() { message[1] = ","; }
        public void AppendWorld() { message[2] = "world"; }

        public string this[int index] { get { return message[index]; } }
    }
}
 
