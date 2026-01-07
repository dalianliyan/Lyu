using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Template.Tests.Exercise
{
    /// <summary>
    /// 模板类型
    /// </summary>
    abstract class InflationRateCalactorBase
    {
        #region 交由子类实现的内容

        public abstract double LastYearIndicator { get; }
        public abstract double ThisYearIndicator { get; }

        #endregion

        /// <summary>
        /// 算法模板
        /// </summary>
        /// <returns></returns>
        public double GetInflationRate()
        {
            return (ThisYearIndicator - LastYearIndicator) / LastYearIndicator;
        }
    }

    class GdpInflationRateCalactor : InflationRateCalactorBase
    {
        public override double LastYearIndicator { get { throw new NotImplementedException(); } }
        public override double ThisYearIndicator { get { throw new NotImplementedException(); } }
    }

    class CpiInflationRateCalactor : InflationRateCalactorBase
    {
        public override double LastYearIndicator { get { throw new NotImplementedException(); } }
        public override double ThisYearIndicator { get { throw new NotImplementedException(); } }
    }

    class PpiInflationRateCalactor : InflationRateCalactorBase
    {
        public override double LastYearIndicator { get { throw new NotImplementedException(); } }
        public override double ThisYearIndicator { get { throw new NotImplementedException(); } }
    }

    /// <summary>
    /// 分析：
    /// 从描述看，尽管计算通胀的指标不同，而且每个指标的计算过程都很复杂，但其算法结构稳定。
    /// 因此，考虑将采用模板定义算法，而将具体指标的计算交给子类完成
    /// </summary>
    [TestClass]
    public class InflationRateFixture
    {

        InflationRateCalactorBase c1;
        InflationRateCalactorBase c2;
        InflationRateCalactorBase c3;

        [TestInitialize]
        public void Initialize()
        {
            c1 = new GdpInflationRateCalactor();
            c2 = new CpiInflationRateCalactor();
            c3 = new PpiInflationRateCalactor();
        }

        [TestMethod]
        public void TestInflationRateCalculatorTemplate()
        {
            ExecuteInflationRateCalactor(c1);
            ExecuteInflationRateCalactor(c2);
            ExecuteInflationRateCalactor(c3);
        }

        void ExecuteInflationRateCalactor(InflationRateCalactorBase calculator)
        {
            if (calculator == null) throw new ArgumentNullException("calculator");
            try
            {
                calculator.GetInflationRate();
            }
            catch(NotImplementedException){}
        }
    }
}
