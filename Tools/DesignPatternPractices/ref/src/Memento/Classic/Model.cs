using System;
namespace MarvellousWorks.PracticalPattern.Memento.Classic
{
    /// <summary>
    /// 为了便于定义抽象状态类型所定义的接口
    /// </summary>
    public interface IState { }

    /// <summary>
    /// 抽象备忘录对象接口
    /// </summary>
    public interface IMemento<T> where T : IState
    {
        T State { get; set; }
    }

    public abstract class MementoBase<T> : IMemento<T>
        where T : IState
    {
        public virtual T State { get; set; }
    }

    /// <summary>
    /// 抽象的发起者对象接口
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// /// <typeparam name="M"></typeparam>
    public interface IOriginator<T, M>
        where T : IState
        where M : IMemento<T>, new()
    {
        IMemento<T> Memento { get;}
    }

    public abstract class OriginatorBase<T, M>
        where T : IState
        where M : IMemento<T>, new()
    {
        /// <summary>
        /// 发起者对象的状态
        /// </summary>
        protected T state;

        /// <summary>
        /// 把状态保存到备忘录，或者从备忘录恢复之前的状态
        /// </summary>
        public virtual IMemento<T> Memento
        {
            get { return new M() { State = this.state }; }
            set
            {
                if (value == null) throw new ArgumentNullException();
                state = value.State;
            }
        }
    }

    /// <summary>
    /// 具体状态类型
    /// </summary>
    public struct Position : IState
    {
        public int X;
        public int Y;
    }

    /// <summary>
    /// 具体备忘录类型
    /// </summary>
    public class Memento : MementoBase<Position> { }

    /// <summary>
    /// 具体发起者类型
    /// </summary>
    public class Originator : OriginatorBase<Position, Memento>
    {
        /// <summary>
        /// 供客户程序使用的非备忘录相关操作
        /// </summary>
        public void UpdateX(int x){state.X = x;}
        public void DecreaseX(){state.X--;}
        public void IncreaseY(){state.Y++;}
        public Position Current { get { return state; } }
    }
}
