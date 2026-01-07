using System;
using System.Diagnostics;
using MarvellousWorks.PracticalPattern.StagePractice.Tests.Properties;

namespace MarvellousWorks.PracticalPattern.StagePractice.Tests.Mock
{
    class AuthenticationHandlerMockBase : IAuthenticationHandler
    {
        protected Func<object> GetValue { get; set; }
        public AuthenticationHandlerMockBase(Func<object> GetValue){this.GetValue = GetValue;}

        public IAuthenticationPolicy Policy { get; set; }
        public IAuthenticationHandler Successor { get; set; }

        public virtual void Handle(CredentialBase credential)
        {
            if (credential == null) throw new ArgumentNullException("credential");
            if (Policy.IsMatch(credential))
            {
                Trace.WriteLine(string.Format(Resources.OutputFormat, "Handler", GetType().Name, GetValue().ToString()));
            }
            if (Successor != null)
                Successor.Handle(credential);
        }
    }

    class LogProcessorCountHandler : AuthenticationHandlerMockBase
    {
        public LogProcessorCountHandler(): base(() => Environment.ProcessorCount){}
    }

    class LogClrVersionHandler : AuthenticationHandlerMockBase
    {
        public LogClrVersionHandler() : base(() => Environment.ProcessorCount) { }
    }

    class LogOsVersionHandler : AuthenticationHandlerMockBase
    {
        public LogOsVersionHandler() : base(() => Environment.OSVersion) { }
    }

}
