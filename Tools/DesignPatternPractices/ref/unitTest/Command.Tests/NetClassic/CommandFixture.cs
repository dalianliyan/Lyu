using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Command.NetClassic;
namespace MarvellousWorks.PracticalPattern.Command.Tests.NetClassic
{
    [TestClass]
    public class CommandFixture
    {
        [TestMethod]
        public void Test()
        {
            var database = new Database()
                               {
                                   OpenConnection = () => { Trace.WriteLine("Open db"); },
                                   ExecuteCommand = () => { return false; },
                                   CloseConnection = () => { Trace.WriteLine("Db closed"); }
                               };
            database.Process();
        }
    }
}