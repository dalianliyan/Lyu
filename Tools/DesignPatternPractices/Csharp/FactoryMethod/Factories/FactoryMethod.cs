using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryMethod
{
    public class FactoryMethod
    {
        static void Main(string[] args)
        {

            Creator productAFactory = new ProductAFactory();
            BaseProduct productA = productAFactory.Create();
            productA.Print();

            Creator productBFactory = new ProductBFactory();
            BaseProduct productB = productBFactory.Create();
            productB.Print();

            Console.Read();
        }
    }
}
