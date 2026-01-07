namespace MarvellousWorks.PracticalPattern.Memento.Tests.Exercise
{
    interface IMemento<T>
        where T : IState
    {
        IState State { get; set; }
    }
}
