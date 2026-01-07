using System;
namespace MarvellousWorks.PracticalPattern.Singleton.Simplest
{
    sealed class Singleton
    {
        Singleton() { }
        [ThreadStatic]
        public static readonly Singleton Instance = new Singleton();
    }
}
