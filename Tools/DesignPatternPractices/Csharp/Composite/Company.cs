using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Composite
{
    /// <summary>
    /// Component
    /// </summary>
    public abstract class Company
    {

        public abstract void Display(int depth);

        public virtual void Add(Company company)
        {

        }

        public virtual void Remove(Company company)
        {

        }
    }

}
