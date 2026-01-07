using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryMethod
{
    public class ProductBFactory: Creator
    {
        public override BaseProduct Create()
        {
            return new ProductA();
        }
    }
}
