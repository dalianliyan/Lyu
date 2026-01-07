using System;
using System.Collections.Generic;
using System.Configuration;
using System.Collections.Specialized;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.FactoryMethod.Generics
{
    public interface IFactory
    {
        IProduct Create();              //  返回默认的 concrete product
        IProduct Create(string name);   //  按照配置的逻辑名称返回指定的concrete product
    }

    public interface IBatchFactory : IFactory
    {
        IEnumerable<IProduct> Create(string name, int quantity);
        IEnumerable<IProduct> Create(int quantity);        
    }

    public class ProductRegistry
    {
        const string SectionName = "concreteProducts";
        static NameValueCollection collection =
            (NameValueCollection)ConfigurationManager.GetSection(SectionName);

        public Type this[string name]
        {
            get { return Type.GetType(collection[name]); }
        }
    }

    public class Factory : IFactory
    {

        public const string DefaultName = "DEFAULT";
        public readonly ProductRegistry productRegistry = new ProductRegistry();

        #region IFactory Members

        public IProduct Create()
        {
            return (IProduct)Activator.CreateInstance(productRegistry[DefaultName]);
        }

        public IProduct Create(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
            return (IProduct)Activator.CreateInstance(productRegistry[name]);
        }

        #endregion
    }

    public class BatchFactory : Factory, IBatchFactory
    {
        #region IBatchFactory Members

        /// <summary>
        /// batch create
        /// </summary>
        /// <param name="quantity"></param>
        /// <returns></returns>
        public IEnumerable<IProduct> Create(int quantity)
        {
            if (quantity <= 0) throw new IndexOutOfRangeException("quantity");
            return InternalCreate<IProduct>(productRegistry[DefaultName], quantity);
        }

        /// <summary>
        /// batch create
        /// </summary>
        /// <param name="name"></param>
        /// <param name="quantity"></param>
        /// <returns></returns>
        public IEnumerable<IProduct> Create(string name, int quantity)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
            return InternalCreate<IProduct>(productRegistry[name], quantity);
        }

        /// <summary>
        /// Helper method
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type"></param>
        /// <param name="quantity"></param>
        /// <returns></returns>
        IEnumerable<T> InternalCreate<T>(Type type, int quantity)
        {
            if (quantity <= 0) throw new IndexOutOfRangeException("quantity");
            if (type == null) throw new ArgumentNullException("type");
            IList<T> result = new List<T>();
            for (int i = 0; i < quantity; i++)
                result.Add((T)Activator.CreateInstance(type));
            return result;
        }

        #endregion
    }
}
