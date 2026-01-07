using MarvellousWorks.PracticalPattern.Proxy.Classic;
namespace Proxy.Host
{
    public class RealSubjectService : ISubject
    {
        #region ISubject Members

        public string Request()
        {
            return RealSubject.Instance.Request();
        }

        #endregion
    }
}
