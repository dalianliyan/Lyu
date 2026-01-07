namespace MarvellousWorks.PracticalPattern.Bridge.Classic
{
    public interface IImpl
    {
        void OperationImpl();
    }

    public interface IAbstraction
    {
        /// <summary>
        /// 这一步很关键
        /// 他将对于N个变换因素的依赖变成对于 “1个抽象 + N-1个变化因素”的依赖关系
        /// 对变化的复杂性作了一次剥离
        /// </summary>
        IImpl Implementor { get; set;}
        void Operation();
    }

    public class ConcreteImplementatorA : IImpl
    {
        public void OperationImpl() { }
    }
    public class ConcreteImplementatorB : IImpl
    {
        public void OperationImpl() { }
    }

    public class RefinedAbstraction : IAbstraction
    {
        /// <summary>
        /// setter方式依赖注入
        /// </summary>
        public IImpl Implementor { get; set; }

        public void Operation()
        {
            // 其他处理
            Implementor.OperationImpl();
        }
    }
}
