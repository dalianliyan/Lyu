using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    public class Vehicle
    {
        public IEnumerable<string> Wheels { get; set; }
        public IEnumerable<string> Lights { get; set; }
    }
}
