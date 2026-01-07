using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryMethod
{

    public enum ProductName
    {
        ProductA,
        ProductB
    }

    class StaticFactoryMethod
    {
        public static IProduct Create(ProductName name)
        {
            if (name == ProductName.ProductA)
            {
                return new ProductA();
            }
            else if (name == ProductName.ProductB)
            {
                return new ProductB();
            }
            else
            {
                throw new Exception("Product Name is invalid.");
            }
        }
    }
}
