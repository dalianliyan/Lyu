using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryMethod
{
    public class ProductAFactory : Creator
    {
        public override BaseProduct Create()
        {
            return new ProductB();
        }
    }
}
