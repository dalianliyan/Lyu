using System;
using System.Diagnostics;
namespace MarvellousWorks.PracticalPattern.Command.NetClassic
{
    public class Database
    {
        #region Commmand
        public Action OpenConnection { get; set; }
        public Func<bool> ExecuteCommand { get; set; }
        public Action CloseConnection { get; set; }
        #endregion

        public void Process()
        {
            OpenConnection();
            Trace.WriteLine(string.Format("Command result : {0}", ExecuteCommand()));
            CloseConnection();
        }
    }
}
