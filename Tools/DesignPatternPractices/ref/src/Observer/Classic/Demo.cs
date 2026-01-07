using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Observer.Classic
{
    /// <summary>
    /// 观察者接口
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IObserver<T>
    {
        void Update(SubjectBase<T> subject);
    }

    /// <summary>
    /// 抽象目标对象
    /// </summary>
    public abstract class SubjectBase<T> 
    {
        /// <summary>
        /// 登记所有需要通知的观察者
        /// </summary>
        protected List<IObserver<T>> observers = new List<IObserver<T>>();
        protected T state;
        public virtual T State{get{return state;}}

        /// <summary>
        /// Attach
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="observer"></param>
        /// <returns></returns>
        public static SubjectBase<T> operator +(SubjectBase<T> subject, IObserver<T> observer)
        {
            subject.observers.Add(observer);
            return subject;
        }
        /// <summary>
        /// Detach
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="observer"></param>
        /// <returns></returns>
        public static SubjectBase<T> operator -(SubjectBase<T> subject, IObserver<T> observer)
        {
            subject.observers.Remove(observer);
            return subject;
        }

        /// <summary>
        /// 更新各观察者
        /// </summary>
        public virtual void Notify()
        {
            observers.ForEach(x=>x.Update(this));
        }

        /// <summary>
        /// 供客户程序对目标对象进行操作的方法
        /// </summary>
        /// <param name="state"></param>
        public virtual void Update(T state)
        {
            this.state = state;
            Notify();   // 触发对外通知
        }
    }
}
