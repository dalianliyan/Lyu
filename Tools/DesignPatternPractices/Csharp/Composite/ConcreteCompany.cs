using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Composite
{
    /// <summary>
    /// Composite
    /// </summary>
    public class ConcreteCompany : Company
    {
        public ConcreteCompany(String name, String product)
        {
            this.Name = name;
            this.Product = product;
            Companies = new List<Company>();
        }

        public string Name { get; set; }
        public string Product { get; set; }
        public List<Company> Companies { get; set; }


        public override void Display(int depth)
        {
            string strBuf = string.Empty;
            for (int i = 0; i < depth; i++)
            {
                strBuf += ("--");
            }

            System.Console.WriteLine(strBuf + this.Name);
            System.Console.WriteLine("The major service is " + this.Product);

            //用visitor模式实现更好
            //我们的组织结构是公司+部门这两层，要把公司到部门的所有信息都打印出来，所以到了子公司这一层，层数就得+2了。

            foreach (Company company in Companies)
            {
                company.Display(depth + 2);
            }
        }


        public override void Add(Company company)
        {
            this.Companies.Add(company);
        }


        public override void Remove(Company company)
        {
            this.Companies.Remove(company);
        }
    }
}
