using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SGJ.Utilities;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            SGJ.Utilities.WordReporter.WriteWord(@"D:\Liyan\Github\CSharpPractice\SGJ.Utilities\bin\Debug\EDRI_ISSUE_REQ_TEMPLATE.docx", TestWord.GetIssueTitleBlock());
        }
    }
}
