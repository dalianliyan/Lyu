using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Singleton
{
    class SingletonSTA
    {
        /// <summary>
        /// 定义一个静态变量来保存类的实例
        /// </summary>
        private static SingletonSTA instance;

        /// <summary>
        /// 定义私有构造函数，使外界不能创建该类实例
        /// </summary>
        private SingletonSTA()
        {
        }

        /// <summary>
        /// 定义公有方法提供一个全局访问点,同时你也可以定义公有属性来提供全局访问点
        /// </summary>
        /// <returns></returns>
        public static SingletonSTA GetInstance()
        {
            // 如果类的实例不存在则创建，否则直接返回if (uniqueInstance == null)
            {
                instance = new SingletonSTA();
            }
            return instance;
        }
    }
}
