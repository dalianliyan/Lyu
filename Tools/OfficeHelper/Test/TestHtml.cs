using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SGJ.Utilities;

namespace Test
{
    class TestHtml
    {
        static void Test(string[] args)
        {
            List<Parameter> lcolPara = new List<Parameter>(){
                new Parameter(){Name = "1", IsSuccess = true, Remark ="xxx"},
                new Parameter(){Name = "2", IsSuccess = false, Remark ="xxx"},
                new Parameter(){Name = "3", IsSuccess = true, Remark ="xxx"},
                new Parameter(){Name = "4", IsSuccess = true, Remark ="xxx"}
            };
            System.IO.File.WriteAllText(@"C:\Users\yli01\Desktop\1.html", HtmlReporter.Generate(lcolPara));
        }
    }
}
