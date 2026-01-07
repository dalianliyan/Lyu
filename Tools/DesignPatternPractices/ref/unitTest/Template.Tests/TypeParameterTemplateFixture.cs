using System;

namespace MarvellousWorks.PracticalPattern.Template.Tests
{
    class TypeParameterTemplateFixture
    {
        interface IVehicle
        {
            void Run(double kilometers);
        }

        /// <summary>
        /// 测速器
        /// </summary>
        /// <typeparam name="T">对于算法中可变内容——数据结构的抽象</typeparam>
        /// <remarks>
        ///     尽管并没有按照经典模板方法设计抽象的IAbstract
        ///     但它同样将算法中可变的部分交给子类实现
        ///     与经典模板方法模式不同的是，它不是交给自己的子类而是泛型类型参数的子类
        /// </remarks>
        class SpeedDetector<T>
            where T : IVehicle, new()
        {
            const int Distance = 100;

            /// <summary>
            /// 算法结构固定
            /// </summary>
            /// <returns></returns>
            public double GetSpeed()
            {
                var start = DateTime.Now;
                var vehicle = new T();
                vehicle.Run(Distance);
                return Distance/(DateTime.Now - start).TotalHours;
            }
        }
    }
}
