using System.Collections.Generic;

namespace MarvellousWorks.PracticalPattern.Composite.Classic
{
    public class Composite : Component
    {
        public Composite()
        {
            base.children = new List<Component>();
        }
    }
}
