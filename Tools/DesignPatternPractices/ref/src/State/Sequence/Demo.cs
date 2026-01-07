namespace MarvellousWorks.PracticalPattern.State.Sequence
{
    /// <summary>
    /// 抽象的状态接口
    /// </summary>
    public interface IState
    {
        void Open(IContext context);
        void Close(IContext context);
        void Query(IContext context);
    }

    public interface IContext
    {
        IState State { get; set; }
        IContext Open();
        IContext Close();
        IContext Query();
    }

    /// <summary>
    /// 抽象的Context对象
    /// </summary>
    public abstract class ContextBase : IContext
    {
        /// <summary>
        /// 实际控制处理的状态对象
        /// </summary>
        public virtual IState State { get; set; }

        public virtual IContext Open() 
        {
            State.Open(this);
            return this;
        }
        public virtual IContext Close() 
        {
            State.Close(this);
            return this;
        }
        public virtual IContext Query()
        {
            State.Query(this);
            return this;
        }
    }
}
