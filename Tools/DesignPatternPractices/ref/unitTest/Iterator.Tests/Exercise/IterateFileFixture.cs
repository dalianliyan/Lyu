using System;
using System.Diagnostics;
using System.Linq;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Iterator.Tests.Exercise
{
    [TestClass]
    public class IterateFileFixture
    {
        const string Path = @".";

        [TestMethod]
        public void TestListModifiedAndNewCreatedFiles()
        {
            Func<DateTime, bool> lessThanThreeDaysHandler = x => (DateTime.Now - x).Days <= 3;
            var files =
                new DirectoryInfo(Path).GetFiles().Where(x =>
                                                         lessThanThreeDaysHandler(x.LastWriteTime) ||
                                                         lessThanThreeDaysHandler(x.CreationTime)
                    );

            //  当前单元测试目录下一定会有些新编译的文件
            Assert.IsTrue(files.Count() > 0);

            //  输出匹配信息
            foreach(var file in files)
            {
                Trace.Write(file.FullName);
                if(lessThanThreeDaysHandler(file.LastWriteTime))
                    Trace.Write("\tM");
                if (lessThanThreeDaysHandler(file.CreationTime))
                    Trace.Write("\tC");
                Trace.WriteLine("");
            }
        }
    }
}
