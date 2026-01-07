using System.Linq;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Template.Classic
{
    /// <summary>
    /// 抽象的模版接口
    /// </summary>
    public interface IAbstract
    {
        int Quantity { get;}
        double Total { get;}
        double Average { get;}
    }

    /// <summary>
    /// 定义了算法梗概的抽象类型
    /// </summary>
    public abstract class AbstractBase : IAbstract
    {
        public abstract int Quantity { get;}
        public abstract double Total { get;}

        /// <summary>
        /// 算法梗概
        /// </summary>
        public virtual double Average { get { return Total / Quantity; } }
    }

    /// <summary>
    /// 具体类型
    /// </summary>
    public class ArrayData : AbstractBase
    {
        double[] data = new double[3] { 1.1, 2.2, 3.3 };
        public override int Quantity { get { return data.Length; } }
        public override double Total{get { return data.Sum(); }}
    }

    /// <summary>
    /// 具体类型
    /// </summary>
    public class GridData : AbstractBase
    {
        IEnumerable<IEnumerable<double>> data = new double[][]{
                                                 new double[]{1.1, 0, 0},
                                                 new double[]{0, 2.2, 0},
                                                 new double[]{0, 0, 3.3}
                                             };
        public override int Quantity { get { return data.Sum(x=>x.Count()); }}
        public override double Total { get { return data.Sum(x=>x.Sum()); } }
    }
}
