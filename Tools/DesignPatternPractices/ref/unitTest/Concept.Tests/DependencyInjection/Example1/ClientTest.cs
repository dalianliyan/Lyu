using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using MarvellousWorks.PracticalPattern.Concept.DependencyInjection.Example1;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.DependencyInjection.Example1
{
    [TestClass()]
    public class ClientTest
    {
        [TestMethod]
        public void Test()
        {
            Client client = new Client();
            Trace.WriteLine(client.GetYear());
        }
    }
}
