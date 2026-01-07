using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarvellousWorks.PracticalPattern.Composite.Classic
{
    public class Leaf : Component
    {
        /// <summary>
        /// 明确声明不支持此类操作
        /// 这里也恰恰体现出《设计模式》一书中Leaf类型经典设计上的不严谨之处
        /// </summary>
        /// <param name="child"></param>
        public override void Add(Component child)
        {
            throw new NotSupportedException();
        }
        public override void Remove(Component child)
        {
            throw new NotSupportedException();
        }
        public override Component this[int index]
        {
            get { throw new NotSupportedException(); }
        }
    }
}
