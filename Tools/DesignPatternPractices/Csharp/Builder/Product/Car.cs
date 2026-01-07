using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    public class Car
    {
        [BuildStep(2, 4)]
        public void AddWheel(){}
        public void AddEngine() { }
        [BuildStep(1)]
        public void AddBody() { }
    }
}
