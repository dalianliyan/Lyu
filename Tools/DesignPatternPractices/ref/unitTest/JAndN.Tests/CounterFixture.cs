using System;
using System.Linq;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests
{
    [TestClass]
    public class CounterFixture
    {
        class InstrumentLogger
        {
            public void Process(int newValue)
            {
                Trace.WriteLine("Log " + newValue);
            }
        }

        class Counter
        {
            int currentValue;

            /// <summary>
            /// namespace System
            /// {
            ///     public delegate void Action<in T>(T obj);
            /// }
            /// </summary>
            public Action<int> OnChanged;

            public static Counter operator ++(Counter c1)
            {
                c1.currentValue++;
                c1.OnChanged(c1.currentValue);
                return c1;
            }
        }

        [TestMethod]
        public void AddCounter()
        {
            var counter = new Counter();
            var logger = new InstrumentLogger();
            counter.OnChanged += logger.Process;    // = Register()
            counter.OnChanged += (x) => { Trace.WriteLine("display new value " + x); };
            counter.OnChanged += (x) => { if (x >= 3) Trace.WriteLine("Alert " + x); };
            Trace.WriteLine("----------- 1 -----------");
            counter++;
            counter++;
            counter++;
            Trace.WriteLine("----------- 2 -----------");
            counter.OnChanged -= logger.Process;    // = Unregister()
            counter++;
        }
    }
}
