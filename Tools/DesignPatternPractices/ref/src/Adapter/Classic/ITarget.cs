namespace MarvellousWorks.PracticalPattern.Adapter.Classic
{
    public interface ITarget
    {
        void Request();
    }

    public class Adaptee
    {
        public void SpecifiedRequest() { }
    }
}
