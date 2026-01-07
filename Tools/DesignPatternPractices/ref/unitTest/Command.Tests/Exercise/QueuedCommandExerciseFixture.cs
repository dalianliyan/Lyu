using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Command.Tests.Exercise
{
    [TestClass]
    public class QueuedCommandExerciseFixture
    {
        List<string> messages = new List<string>{"hello","command", "pattern"};
        
        [TestMethod]
        public void TestQueuedSerializedCommand()
        {
            //  定义保存命令对象的队列
            //  命令对象的抽象形式为返回结果为string的空参数方法 Func<string>
            var queue = new Queue<Func<string>>();

            //  将命令对象入队列
            messages.ForEach(x=>queue.Enqueue(() => x));

            //  逐个验证出队列后各命令对象是否次序和执行功能正常
            var i = 0;
            while(queue.Count > 0)
            {
                var command = queue.Dequeue();
                Trace.WriteLine(command());
                Assert.AreEqual<string>(messages[i], command());
                i++;
            }
            
        }
    }
}
