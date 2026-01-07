using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder
{
    [AttributeUsage(AttributeTargets.Method,AllowMultiple = false)]
    public sealed class BuildStepAttribute:Attribute
    {
        private int sequence;        
        public int Sequence { get { return this.sequence; } }

        private int times;
        public int Times { get { return this.times; } }
    }
}
