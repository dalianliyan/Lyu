using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    class Program
    {
        static void Main(string[] args)
        {
            var maker = new VehicleMaker();
            maker.Builder = new CarBuilder();
            maker.Construct();
            maker.Builder = new BicycleBuilder();
            maker.Construct();
        }
    }
}
