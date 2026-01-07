namespace MarvellousWorks.PracticalPattern.Adapter.Classic.ClassicObject
{
    public class Adapter : ITarget
    {
        public Adapter(Adaptee adaptee)
        {
            Adaptee = adaptee;
        }

        public Adaptee Adaptee{ get; private set;}    // Adaptee对象

        public void Request()
        {
            // 其他操作
            Adaptee.SpecifiedRequest();    // 调用Adaptee
        }
    }
}
