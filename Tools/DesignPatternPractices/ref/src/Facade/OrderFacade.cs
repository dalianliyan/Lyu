using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Facade
{
    public class User{}
    public class Product{}

    /// <summary>
    /// 订单操作外观类
    /// </summary>
    public class OrderFacade
    {
        /// <summary>
        /// 下单
        /// </summary>
        /// <param name="detail">购物内容明细</param>
        /// <param name="user">用户信息</param>
        /// <returns>应付金额</returns>
        /// <remarks>如果库存不足抛出异常</remarks>
        public double Add(IEnumerable<KeyValuePair<Product, double>> detail, User user)
        {
            throw new NotImplementedException();
        }
    }
}
