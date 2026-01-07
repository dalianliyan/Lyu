using System;
using System.Diagnostics;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;
namespace MarvellousWorks.PracticalPattern.Proxy.Remote
{
    public class MessageTracer : IDispatchMessageInspector, IClientMessageInspector
    {
        #region IDispatchMessageInspector Members

        public object AfterReceiveRequest(ref Message request, 
            IClientChannel channel, InstanceContext instanceContext)
        {
            Trace.WriteLine("Service side : afater receive request");
            return null;
        }

        public void BeforeSendReply(ref Message reply, object correlationState)
        {
            Trace.WriteLine("Service side : before send reply");
        }

        #endregion

        #region IClientMessageInspector Members

        public void AfterReceiveReply(ref Message reply, object correlationState)
        {
            Trace.WriteLine("Client side : after receive reply");
        }

        public object BeforeSendRequest(ref Message request, IClientChannel channel)
        {
            Trace.WriteLine("Client side : before send request");
            return null;
        }

        #endregion
    }
}
