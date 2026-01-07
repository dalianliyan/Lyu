using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryMethod
{
    public class ProductB : BaseProduct
    {

        public override void Print()
        {
            Console.WriteLine("Product B has been created!");
        }
    }
}
