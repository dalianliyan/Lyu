using System;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.StagePractice.Tests
{
    [TestClass]
    public class FileSystemWatcherFixture
    {
        const string TargetFileName = "hello.txt";

        void OnChanged(object target, FileSystemEventArgs args)
        {
            Console.WriteLine(string.Format("File {0} {1}", args.FullPath, args.ChangeType));
        }

        [TestMethod]
        public void TestChangeFile()
        {
            var watcher = new FileSystemWatcher();
            watcher.Path = Environment.CurrentDirectory;
            watcher.NotifyFilter = NotifyFilters.LastWrite;
            watcher.Filter = TargetFileName;
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.EnableRaisingEvents = true;

            Console.WriteLine("Current content of {0} is \n{1}", TargetFileName, File.ReadAllText(TargetFileName));
            File.AppendAllText(TargetFileName, "\nchanged");
            Thread.Sleep(1000);
            Console.WriteLine("Content after modify of {0} is \n{1}", TargetFileName, File.ReadAllText(TargetFileName));
        }
    }
}
