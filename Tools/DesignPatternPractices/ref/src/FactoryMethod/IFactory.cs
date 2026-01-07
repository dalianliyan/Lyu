using System;
using System.Collections.Generic;
using System.Data.Linq;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.FactoryMethod
{
    /// <summary>
    /// 工厂接口定义
    /// </summary>
    /// <remarks>
    ///     TTarget : abstract product type
    ///     TSource:  concrete product type
    /// </remarks>
    public interface IFactory
    {
        #region config and register type mapping

        /// <summary>
        /// 如果需要同时加载配置文件中定义的映射关系，可以按照SRP的原则定义独立的配置类型。
        /// 由该配置类型调用这两个接口为Factory加载配置信息
        /// </summary>

        IFactory RegisterType<TTarget, TSource>();  // fluent interface
        IFactory RegisterType<TTarget, TSource>(string name);   // fluent interface

        #endregion

        #region factory method

        TTarget Create<TTarget>();
        TTarget Create<TTarget>(string name);

        #endregion
    }

    public sealed class TypeRegistry
    {
        /// <summary>
        /// default name in type mappings
        /// </summary>
        readonly string DefaultName = Guid.NewGuid().ToString();

        /// <summary>
        /// Type        :   TTarget, abstract product type
        /// IDictionary<string ,Type>
        ///     string  :   name
        ///     Type    :   TSource, concrete product type
        /// </summary>
        IDictionary<Type, IDictionary<string, Type>> registry =
            new Dictionary<Type, IDictionary<string, Type>>();

        public void RegisterType(Type targetType, Type sourceType)
        {
            RegisterType(targetType, sourceType, DefaultName);
        }

        public void RegisterType(Type targetType, Type sourceType, string name)
        {
            if(targetType == null) throw new ArgumentNullException("targetType");
            if(sourceType == null) throw new ArgumentNullException("sourceType");
            if(string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");

            IDictionary<string, Type> subDictionary;
            if(!registry.TryGetValue(targetType, out subDictionary))
            {
                subDictionary = new Dictionary<string, Type>();
                subDictionary.Add(name, sourceType);
                registry.Add(targetType, subDictionary);
            }
            else
            {
                if(subDictionary.ContainsKey(name))
                    throw new DuplicateKeyException(name);
                subDictionary.Add(name, sourceType);
            }
        }

        public Type this[Type targetType, string name]
        {
            get
            {
                if (targetType == null) throw new ArgumentNullException("targetType");
                if (string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
                if (registry.Count() == 0)
                    return null;

                return 
                    // filter by target type
                    (registry.Where(x => x.Key == targetType)).FirstOrDefault().Value
                    // filter by type mapping name
                    .Where(x => string.Equals(name, x.Key)).FirstOrDefault().Value;
            }
        }

        public Type this[Type targetType]
        {
            get { return this[targetType, DefaultName]; }
        }
    }

    public class Factory : IFactory
    {
        protected TypeRegistry registry = new TypeRegistry();

        #region IFactory Members

        public IFactory RegisterType<TTarget, TSource>()
        {
            registry.RegisterType(typeof(TTarget), typeof(TSource));
            return this;
        }

        public IFactory RegisterType<TTarget, TSource>(string name)
        {
            registry.RegisterType(typeof(TTarget), typeof(TSource), name);
            return this;
        }

        public TTarget Create<TTarget>()
        {
            return (TTarget)Activator.CreateInstance(registry[typeof(TTarget)]);
        }

        public TTarget Create<TTarget>(string name)
        {
            return (TTarget)Activator.CreateInstance(registry[typeof(TTarget), name]);
        }

        #endregion
    }

}
