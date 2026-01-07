using System;
namespace MarvellousWorks.PracticalPattern.Concept.Generics
{
    #region static impl.
    //public static class RawGenericFactory
    //{
    //    public static T Create<T>(string typeName)
    //    {
    //        return (T)Activator.CreateInstance(Type.GetType(typeName));
    //    }
    //}
    #endregion

    /// <summary>
    /// 原始的泛型工厂类型
    /// </summary>
    /// <typeparam name="T">构造结果的目标类型，一般T是抽象类型或接口</typeparam>
    public class RawGenericFactory<T>
    {
        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="typeName">类型名称</param>
        /// <returns>构造实例结果</returns>
        public T Create(string typeName)
        {
            if(string.IsNullOrEmpty(typeName)) throw new ArgumentNullException("typeName");
            return (T)Activator.CreateInstance(Type.GetType(typeName));
        }
    }
}
