namespace MarvellousWorks.PracticalPattern.OO.SRP
{
    interface IReader<TData>
    {
        TData Read();
    }
    interface IWriter<TData>
    {
        void Write(TData graph);
    }

    interface IDataFormatter<TSource, TTarget>
    {
        TTarget Format(TSource source);
    }

    interface IDataHandler<TSource, TTarget> : 
        IReader<TSource>, IWriter<TTarget>, IDataFormatter <TSource, TTarget>{}

    //interface IDataHandler<TSource, TTarget>
    //{
    //    TSource Read();
    //    TTarget Format(TSource source);
    //    void Write(TTarget graph);
    //}
}
