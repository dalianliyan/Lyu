using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace Builder.Tag
{
    public class BuildStep
    {
        public MethodInfo Method { get; set; }
        public int Times { get; set; }
        public int Sequence { get; set; }
    }
}
