using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    [TestClass]
    public class GenericsSupportFixture
    {
        [TestMethod]
        public void TestNonGenericCollection()
        {
            var collection = new Hashtable();
            collection.Add(1, 2);
        }

        [TestMethod]
        public void TestGenericCollection()
        {
            var collection = new Dictionary<int, string>();
            collection.Add(1, "2");
        }
    }

    //public abstract class Person { }
    //public class OldMan : Person { }
    //public class Child : Person { }

    ///// <summary>
    ///// 但很可惜，类型参数间的与或关系目前只是开发人员一个美好的愿望
    ///// </summary>
    ///// <typeparam name="T"></typeparam>
    //public class TicketFreeCommandA<T>
    //    where T : OldMan || Child
    //{
    //}
}
