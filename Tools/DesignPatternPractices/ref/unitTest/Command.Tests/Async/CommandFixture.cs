using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Command.Tests.Async
{
    [TestClass]
    public class CommandFixture
    {
        class AsyncObject
        {
            public void Execute(IEnumerable<int> data)
            {
                if (data == null) throw new ArgumentNullException("data");
                Action<IEnumerable<int>> callback = 
                    (x) =>
                    {
                        if (x == null) throw new ArgumentNullException("x");
                        foreach (var item in x)
                            Thread.Sleep(200);
                        Trace.WriteLine("real finished");
                    };
                callback.BeginInvoke(data, callback.EndInvoke, null);
                Trace.WriteLine("fake finished");
            }
        }

        [TestMethod]
        public void Test()
        {
            var target = new AsyncObject();
            target.Execute(new int[]{3, 5, 2, 7, 1});
            Thread.Sleep(3000); // 须有此行代码，否则单元测试就退出，无法看到实际执行结束
        }
    }
}
