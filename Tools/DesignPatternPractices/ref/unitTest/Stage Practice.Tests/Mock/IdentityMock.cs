using System.Security.Principal;
namespace MarvellousWorks.PracticalPattern.StagePractice.Tests.Mock
{
    class IdentityMock : IIdentity
    {
        public bool IsAuthenticated { get; set; }
        public string Name { get; set; }
        public string AuthenticationType { get; set; }

        public string State { get; set; }

        public override string ToString()
        {
            return string.Format("\n\tIsAuthencated : {0}\n\t Name : {1}\n\t Credential Type : {2}\n\t State : {3}",
                                 IsAuthenticated, Name, AuthenticationType, State);
        }
    }
}
