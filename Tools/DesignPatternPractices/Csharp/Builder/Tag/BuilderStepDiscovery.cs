using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builder.Tag
{
    public class BuilderStepDiscovery
    {
        static IDictionary<Type, IEnumerable<BuildStep>> cache =
            new Dictionary<Type, IEnumerable<BuildStep>>();
        static IList<Type> errorCache = new List<Type>();
    }
}
