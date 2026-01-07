using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    [TestClass]
    public class NonGenericCollectionConvertFixture
    {
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestConvertNonGenericDictionary()
        {
            var e1 = new DictionaryEntry("2", "second");
            var e2 = new KeyValuePair<string, string>(e1.Key.ToString(), e1.Value.ToString());

            var gd = new Dictionary<string, string>();
            gd.Add("a", "first");
            gd.Add("b", "second");
            
            var ht = new Hashtable();
            ht.Add("a", "first");
            ht.Add("b", "second");

            ht.AsQueryable();   // here
            var result = ht.Cast<KeyValuePair<string, string>>();
            foreach(var item in result)
                Trace.WriteLine(item);
        }
    }
}
