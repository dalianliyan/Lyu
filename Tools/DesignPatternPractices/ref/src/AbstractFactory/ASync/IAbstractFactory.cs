using System;
namespace MarvellousWorks.PracticalPattern.AbstractFactory.Async
{
    public interface IFactory
    {
        IProduct Create();
    }
    public interface IFactoryWithNotifier : IFactory
    {
        void Create(Action<IProduct> callback);
    }
}
