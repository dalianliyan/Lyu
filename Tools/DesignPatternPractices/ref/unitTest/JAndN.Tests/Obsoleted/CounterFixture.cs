using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests.Obsoleted
{
    [TestClass]
    public class CounterFixture
    {
        interface IInstrumentTarget
        {
            void Process(int newValue);
        }

        class InstrumentLogger : IInstrumentTarget
        {
            public void Process(int newValue)
            {
                Trace.WriteLine("Log " + newValue);
            }
        }

        class RuleEngine : IInstrumentTarget
        {
            const int AlertValue = 3;
            public void Process(int newValue)
            {
                if (newValue >= AlertValue)
                    Trace.WriteLine("Alert " + newValue);
            }
        }

        class InstrumentWindowUI : IInstrumentTarget
        {
            public void Process(int newValue)
            {
                Trace.WriteLine("display new value " + newValue);
            }
        }

        class Counter
        {
            int currentValue;
            List<IInstrumentTarget> instrumentTargets = new List<IInstrumentTarget>();

            public static Counter operator ++(Counter c1)
            {
                c1.currentValue++;
                c1.instrumentTargets.ForEach(x => x.Process(c1.currentValue));
                return c1;
            }

            public Counter Register(IInstrumentTarget target)
            {
                if (target == null)
                    throw new ArgumentNullException("target");
                instrumentTargets.Add(target);
                return this;
            }

            public Counter Unregister(IInstrumentTarget target)
            {
                if (target == null)
                    throw new ArgumentNullException("target");
                instrumentTargets.Remove(target);
                return this;
            }
        }

        [TestMethod]
        public void AddCounter()
        {
            var logger = new InstrumentLogger();
            var counter = new Counter()
                .Register(logger)
                .Register(new RuleEngine())
                .Register(new InstrumentWindowUI());
            Trace.WriteLine("----------- 1 -----------");
            counter++;
            counter++;
            counter++;
            Trace.WriteLine("----------- 2 -----------");
            counter.Unregister(logger);
            counter++;
        }
    }
}
