using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryMethod
{
    public class ProductA : BaseProduct
    {

        public override void Print()
        {
            Console.WriteLine("Product A has been created!");
        }
    }
}
