namespace MarvellousWorks.PracticalPattern.Adapter.Classic.ClassicClass
{
    public class Adapter : Adaptee, ITarget
    {
        public void Request()
        {
            // 其他操作
            base.SpecifiedRequest();    // 调用Adaptee
        }
    }
}
