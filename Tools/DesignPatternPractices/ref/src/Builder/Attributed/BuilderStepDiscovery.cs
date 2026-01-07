using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.Builder.Attributed
{
    public class BuildStep
    {
        public MethodInfo Method { get; set; }
        public int Times { get; set; }
        public int Sequence { get; set; }
    }

    /// <summary>
    /// 通过反射获得某个类型相关BuildPart()步骤指导信息的工具类型
    /// </summary>
    public class BuilderStepDiscovery
    {
        /// <summary>
        /// 缓冲已经解析过的类型信息
        /// </summary>
        static IDictionary<Type, IEnumerable<BuildStep>> cache =
            new Dictionary<Type, IEnumerable<BuildStep>>();

        /// <summary>
        /// 登记那些已经认定没有Build Step 属性的类型
        /// </summary>
        static IList<Type> errorCache = new List<Type>();

        /// <summary>
        /// 借助反射获得类型 T 所需执行BuildPart()的自动发现机制
        /// </summary>
        /// <returns></returns>
        public IEnumerable<BuildStep> DiscoveryBuildSteps(Type type)
        {
            if (type == null) throw new ArgumentNullException("type");
            if (errorCache.Contains(type)) return null;
            if (!cache.ContainsKey(type))
            {
                var aType = typeof(BuildStepAttribute);
                var methods = from item in
                    (from method in type.GetMethods()
                     where method.IsDefined(aType, false)
                     select new
                     {
                         M = method,
                         A = (BuildStepAttribute)method.GetCustomAttributes(aType, false).First()
                     }
                     )orderby item.A.Sequence
                              select new BuildStep
                              {
                                  Method = item.M,
                                  Times = item.A.Times,
                                  Sequence = item.A.Sequence
                              };
                if (methods.Count() == 0)
                {
                    errorCache.Add(type); // register invalidate type
                    return null;
                }
                else
                {
                    cache.Add(type, methods); // register validate type
                    return methods;
                }
            }
            else
                return cache[type];
        }
    }
}
