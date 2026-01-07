using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Composite
{
    class Program
    {
        static void Main(string[] args)
        {
            // 测试类（Client）：三个公司，阿里巴巴、淘宝、阿里云，
            // 其中，淘宝和阿里云是阿里巴巴的子公司，
            // 同时，每个公司又有自己的研发部门和人资部门。
            Company alibaba = new ConcreteCompany("Alibaba", "B2B");
            Company taobao = new ConcreteCompany("Taobao", "C2C");
            Company aliyun = new ConcreteCompany("Aliyun", "Cloud Compution");

            Company aliHrDepartment = new Department("Alibaba HR Department");
            Company aliResearchDepartment = new Department("Alibaba RESEARCH Department");

            Company taobaoHrDepartment = new Department("Taobao HR Department");
            Company taobaoResearchDepartment = new Department("Taobao RESEARCH Department");

            Company aliyunHrDepartment = new Department("Aliyun HR Department");
            Company aliyunResearchDepartment = new Department("Aliyun RESEARCH Department");

            alibaba.Add(taobao);
            alibaba.Add(aliyun);
            alibaba.Add(aliHrDepartment);
            alibaba.Add(aliResearchDepartment);

            taobao.Add(taobaoHrDepartment);
            taobao.Add(taobaoResearchDepartment);

            aliyun.Add(aliyunHrDepartment);
            aliyun.Add(aliyunResearchDepartment);

            alibaba.Display(0);
            System.Console.WriteLine("\n~~~~~~~~~~~~~~~~~~~~~~~~\n");
            taobao.Display(0);
            System.Console.WriteLine("\n~~~~~~~~~~~~~~~~~~~~~~~~\n");
            aliyun.Display(0);
            Console.ReadKey();
        }
    }
}
