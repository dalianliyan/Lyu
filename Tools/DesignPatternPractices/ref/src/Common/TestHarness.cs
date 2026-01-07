using System;
using System.Collections.Generic;
using System.Linq;

using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.Common
{
    /// <summary>
    /// Some utility extensions on unit test.
    /// </summary>
    public static class TestHarness
    {
        /// <summary>
        /// run a few delegates and wait untill all to complete
        /// </summary>
        public static void Parallel(params ThreadStart[] actions)
        {
            Thread[] threads = actions.Select(a => new Thread(a)).ToArray();
            threads.ForEach(t => t.Start());
            threads.ForEach(t => t.Join());
        }
    }
}
