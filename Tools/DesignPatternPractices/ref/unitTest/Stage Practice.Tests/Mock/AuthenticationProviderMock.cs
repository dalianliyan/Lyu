using System;
using System.Security.Principal;
using MarvellousWorks.PracticalPattern.StagePractice.Tests.Usb;
namespace MarvellousWorks.PracticalPattern.StagePractice.Tests.Mock
{
    abstract class AuthenticationProviderMockBase : IAuthenticationProvider
    {
        public abstract string AuthenticationType { get; }
        public abstract Type CredentialType { get; }

        public const string Name = "mock";
        public const string State = "provider";
        
        public virtual IIdentity Verify(CredentialBase credential)
        {
            if((credential == null) || (credential.GetType() != CredentialType))
                throw new ArgumentException("credential");
            return new IdentityMock()
            {
                AuthenticationType = AuthenticationType,
                Name = Name,
                State = State,
                IsAuthenticated = true
            };
        }
    }

    class UsbKeyProvider : AuthenticationProviderMockBase
    {
        public override string AuthenticationType { get { return "usb"; } }
        public override Type CredentialType { get { return typeof(UsbKeyCredential); } }

        public override IIdentity Verify(CredentialBase credential)
        {
            var adapter = new UsbAdapter();
            if(!adapter.IsOpen)
                adapter.Open();
            if(credential.GetType() != adapter.GetCredential(UsbAdapter.Pin).GetType())
                throw new NotSupportedException();
            var result = base.Verify(credential);
            adapter.Close();
            return result;
        }
    }

    class WindowsProvider : AuthenticationProviderMockBase
    {
        public override string AuthenticationType { get { return "win"; } }
        public override Type CredentialType { get { return typeof(WindowsCredential); } }
    }

    class UserNameProvider : AuthenticationProviderMockBase
    {
        public override string AuthenticationType { get { return "userName"; } }
        public override Type CredentialType { get { return typeof(UserNameCredential); } }
    }

    class CustomsProvider : AuthenticationProviderMockBase
    {
        public override string AuthenticationType { get { return "customs"; } }
        public override Type CredentialType { get { return typeof(UserNameCredential); } }
    }

}
