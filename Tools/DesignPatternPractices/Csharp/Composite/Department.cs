using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Composite
{
    /// <summary>
    /// Leaf
    /// </summary>
    public class Department : Company
    {


        public string Name { get; set; }


        public Department(String name)
        {
            this.Name = name;
        }


        public override void Display(int depth)
        {

            string sb = string.Empty;
            for (int i = 0; i < depth; i++)
            {
                sb += "--";
            }
            System.Console.WriteLine(sb + this.Name);
        }


        public override void Add(Company company) { }

        public override void Remove(Company company) { }

    }
}
