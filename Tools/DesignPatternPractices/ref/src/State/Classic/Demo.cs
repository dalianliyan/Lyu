using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.State.Classic
{
    /// <summary>
    /// 抽象的状态接口
    /// </summary>
    public interface IState
    {
        void Open();
        void Close();
        void Query();
    }

    /// <summary>
    /// 抽象的Context对象
    /// </summary>
    public abstract class ContextBase
    {
        /// <summary>
        /// 实际控制处理的状态对象
        /// </summary>
        public virtual IState State{get; set;}

        #region state object controlled logic
        public virtual void Open() { State.Open(); }
        public virtual void Close() { State.Close(); }
        public virtual void Query() { State.Query(); }
        #endregion
    }
}
