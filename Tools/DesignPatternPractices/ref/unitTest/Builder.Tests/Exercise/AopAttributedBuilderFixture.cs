using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using B = MarvellousWorks.PracticalPattern.Builder.Attributed;
namespace MarvellousWorks.PracticalPattern.Builder.Tests.Exercise
{
    /// <summary>
    /// BuildUp()适用的EventArgs
    /// </summary>
    /// <typeparam name="T"></typeparam>
    class BuilderEventArgs<T> : EventArgs
    {
        /// <summary>
        /// 还没有进行装配的对象实例
        /// </summary>
        public T Target { get; set; }
    }

    /// <summary>
    /// BuildPart()适用的EventArgs
    /// </summary>
    /// <typeparam name="T"></typeparam>
    class BuilderStepEventArgs<T> : BuilderEventArgs<T>
    {
        /// <summary>
        /// 该BuildStep的操作信息
        /// </summary>
        public B.BuildStep Step { get; set; }

        /// <summary>
        /// 该BuildStep正在执行第几轮循环
        /// </summary>
        public int CurrentLoop { get; set; }
    }

    class Builder<T>  where T : class, new()
    {
        public event EventHandler<BuilderEventArgs<T>> BeforeBuildUp;
        public event EventHandler<BuilderEventArgs<T>> AfterBuildUp;

        public event EventHandler<BuilderStepEventArgs<T>> BeforeBuildStepExecuting;
        public event EventHandler<BuilderStepEventArgs<T>> AfterBuildStepExecuted;

        public virtual T BuildUp()
        {
            var steps = new B.BuilderStepDiscovery().DiscoveryBuildSteps(typeof(T));
            if (steps == null) return new T(); // 没有BuildPart步骤，退化为Factory模式
            var target = new T();

            //  BuildUp() 执行前的切入点
            if (BeforeBuildUp != null)
                BeforeBuildUp(this, new BuilderEventArgs<T>() { Target = target });

            foreach (var step in steps)
                for (var i = 0; i < step.Times; i++)
                {
                    //  BuildPart() 执行前的切入点
                    if (BeforeBuildStepExecuting != null)
                        BeforeBuildStepExecuting(this,
                                                 new BuilderStepEventArgs<T>()
                                                 {
                                                     Target = target,
                                                     Step = step,
                                                     CurrentLoop = i
                                                 });

                    step.Method.Invoke(target, null);

                    //  BuildPart() 执行后的切入点
                    if (AfterBuildStepExecuted != null)
                        AfterBuildStepExecuted(this,
                                                 new BuilderStepEventArgs<T>()
                                                 {
                                                     Target = target,
                                                     Step = step,
                                                     CurrentLoop = i
                                                 });
                }

            //  BuildUp() 执行后的切入点
            if (AfterBuildUp != null)
                AfterBuildUp(this, new BuilderEventArgs<T>() { Target = target });

            return target;
        }
    }

    /// <summary>
    /// 这里模拟一个可能制造出次品的汽车对象
    /// 每个部件都可能在装配过程中出现错误
    /// </summary>
    class Car
    {
        public const string WheelItem = "wheel";
        public const string EngineItem = "engine";
        public const string BodyItem = "body";

        const int WheelDefectiveCycle = 9;
        const int BodyDefectiveCycle = 11;

        static int wheelDefectiveIndexer;
        static int bodyDefectiveIndexer;

        public IList<string> Parts { get; private set; }
        public Car() { Parts = new List<string>(); }

        /// <summary>
        /// 为汽车添加轮胎
        /// </summary>
        /// <remarks>
        ///     Attributed Builder第二个执行的Setp
        ///     执行4次，即为每辆汽车装配增加4个轮胎
        /// </remarks>
        [B.BuildStep(2, 4)]
        public void AddWheel()
        {
            wheelDefectiveIndexer++;

            //  模拟会出现装配错误的情况
            if (wheelDefectiveIndexer % WheelDefectiveCycle != 0)
                Parts.Add(WheelItem);
        }

        /// <summary>
        /// 为汽车装配引擎
        /// </summary>
        /// <remarks>
        ///     没有通过Attribute标注的内容，因此不会被Attributed Builder执行
        /// </remarks>
        public void AddEngine()
        {
            Trace.WriteLine(EngineItem);
            Parts.Add(EngineItem);
        }

        /// <summary>
        /// 为汽车装配车身
        /// </summary>
        /// <remarks>
        ///     Attributed Builder第一个执行的Setp
        ///     执行1次，即为每辆汽车装配增加1个车身
        /// </remarks>
        [B.BuildStep(1)]
        public void AddBody()
        {
            bodyDefectiveIndexer++;

            //  模拟会出现装配错误的情况
            if (bodyDefectiveIndexer % BodyDefectiveCycle != 0)
                Parts.Add(BodyItem);
        }
    }

    [TestClass]
    public class AopAttributedBuilderFixture
    {
        int totalBuildUpCars;   //  装配车辆总数
        int qualifiedCars;      //  装配合格车辆总数
        int defectiveCars;      //  装配次品车辆总数
        int partsBeforeBuild;   //  当前BuildPart()执行前已经装配合格的零件数量

        Builder<Car> builder;

        [TestInitialize]
        public void Initialize()
        {
            totalBuildUpCars = 0;
            qualifiedCars = 0;
            defectiveCars = 0;

            builder = new Builder<Car>();

            //  BuildUp()执行前记录总共装配车辆数量
            builder.BeforeBuildUp += (x, y) => { totalBuildUpCars++; };

            //  BuildUp()执行后记录装配的产品是否合格
            builder.AfterBuildUp +=
                (x, y) =>
                    {
                        //  合格的车辆是4个轮子一个车身
                        if ((y.Target.Parts.Where(part => string.Equals(part, Car.WheelItem)).Count() == 4)
                            && (y.Target.Parts.Where(part => string.Equals(part, Car.BodyItem)).Count() == 1))
                        {
                            qualifiedCars++;
                            Trace.WriteLine("qualified\n");
                        }
                        else
                        {
                            defectiveCars++;
                            Trace.WriteLine("defective\n");
                        }
                    };

            //  BuildPart()执行前，记录当前装配的工序信息
            //  同时登记之前已经合格完成装配的零件数量
            builder.BeforeBuildStepExecuting +=
                (x, y) =>
                    {
                        partsBeforeBuild = y.Target.Parts.Count;
                        if (y.Step.Times > 1)
                            Trace.Write(string.Format("step {0} sequence\t[{1}] current loop [{2}/{3}]\t",
                                y.Step.Method.Name,
                                y.Step.Sequence,
                                y.CurrentLoop + 1,
                                y.Step.Times));
                        else
                            Trace.Write(string.Format("step {0} sequence\t[{1}] \t",
                                                      y.Step.Method.Name,
                                                      y.Step.Sequence));
                    };

            //  通过对比每个BuildPart()过程前后实际已装配的零件数量，可以记录缺陷具体出现在哪个步骤
            builder.AfterBuildStepExecuted +=
                (x, y) =>
                    {
                        if (y.Target.Parts.Count > partsBeforeBuild)
                            Trace.WriteLine("qualified");
                        else
                            Trace.WriteLine("defective");
                    };
        }

        /// <summary>
        /// 单元测试中对于整车和每个部件的检测工作通过AOP手段嵌入BuildUp()过程中
        /// 合格车辆的标准是每个整车和每个部件（1个车身和4个轮子）都合格
        /// </summary>
        [TestMethod]
        public void TestAopDuringBuildUp()
        {
            for (var i = 0; i < 7; i++ )
                builder.BuildUp();

            Trace.WriteLine("\n\n ---------------------------------------");
            Trace.WriteLine("total build up " + totalBuildUpCars);
            Trace.WriteLine("qualified " + qualifiedCars);
            Trace.WriteLine("defective " + defectiveCars);
        }
    }
}
