using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Command.Packaging;
namespace MarvellousWorks.PracticalPattern.Command.Tests.Packaging
{
    [TestClass]
    public class CommandFixture
    {
        [TestMethod]
        public void Test()
        {
            var dt = new DbOperDT();
            var ra = dt.Insert("a");
            dt.Update(ra, "a1").Update(ra, "a11").Update(ra, "a111");
            var rb = dt.Insert("b");
            dt.Update(rb, "b2").Update(rb, "b2222").Update(rb, "b22");
            dt.Delete(rb);
            var rc = -1;   // 既有的一个记录 
            dt.Update(rc, "c3").Update(rc, "c33").Delete(rc);
            var rd = -2;
            dt.Update(rd, "d444").Update(rd, "d4");

            dt.Update(ra, "ax").Update(rd, "dy");

            Trace.WriteLine("\n before merge");
            dt.Profile();

            Trace.WriteLine("\n\n after merge");
            dt.Merge();
            dt.Profile();
        }
    }
}