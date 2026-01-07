using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Observer.ObserverCollection.Simple;
namespace MarvellousWorks.PracticalPattern.Observer.Tests.ObserverCollection.Simple
{
    [TestClass]
    public class ObserverFixture
    {
        string key = "hello";
        string value = "world";

        void Validate(object sender, DictionaryEventArgs<string, string> args)
        {
            Assert.IsNotNull(sender);
            Assert.IsTrue(sender is ObserverableDictionary<string, string>);
            Assert.IsNotNull(args);
            Assert.AreEqual<string>(key, args.Key);
            Assert.AreEqual<string>(value, args.Value);
        }

        [TestMethod]
        public void Test()
        {
            var dictionary = new ObserverableDictionary<string, string>();
            dictionary.NewItemAdded += this.Validate;
            dictionary.Add(key, value);
        }
    }
}
