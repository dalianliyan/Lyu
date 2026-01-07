namespace MarvellousWorks.PracticalPattern.Mediator.Classic
{
    /// <summary>
    /// 协同对象接口
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IColleague<T>
    {
        T Data { get; set; }
        IMediator<T> Mediator { get; set; }
    }

    public abstract class ColleagueBase<T> : IColleague<T>
    {
        public virtual T Data { get; set; }
        public virtual IMediator<T> Mediator { get; set; }
    }
}
