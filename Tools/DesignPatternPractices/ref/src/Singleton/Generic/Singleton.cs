using System;
using System.Reflection;
namespace MarvellousWorks.PracticalPattern.Singleton.Generic
{
    /// <summary>
    /// 提供对封闭构造方法类型的Signelton包装
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class Singleton<T>
        where T : class
    {
        private Singleton()
        {
        }

        public static readonly T Instance =
            (T)typeof(T).
                GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new Type[0], null).
                Invoke(null);
    }
}