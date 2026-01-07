using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Iterator.Pure;
namespace MarvellousWorks.PracticalPattern.Iterator.Tests.Pure
{
    [TestClass]
    public class IteratorFixture
    {
        Repository repository;

        [TestInitialize]
        public void Initialize()
        {
            repository = new Repository()
            {
                ArrayData = new int[] { 1, 2 },
                CheckList = new Dictionary<int, int>() { { 1, 3 }, { 2, 4 }, { 3, 5 } },
                Root = new TreeNode<int>()
                {
                    Data = 6,
                    children = new List<TreeNode<int>>()
                    {
                        new TreeNode<int>() {Data = 7},
                        new TreeNode<int>() {Data = 8},
                        new TreeNode<int>() {Data = 9},
                        new TreeNode<int>()
                        {
                            Data = 10,
                            children = new List<TreeNode<int>>()
                            {
                                new TreeNode<int>(){Data = 11},
                                new TreeNode<int>(){Data = 12},
                            }
                        }
                    }
                }
            };
        }

        [TestMethod]
        public void LinkIterator()
        {
            Assert.AreEqual<int>(15, repository.GetAll().Count());
            int count = 0;
            repository.GetAll().ToList().ForEach(x => Assert.AreEqual<int>(++count, x));
        }

        [TestMethod]
        public void IterateWithFilter()
        {
            int count = 2;
            repository.GetAll().Where(x => x % 2 == 0).ToList().ForEach(x =>
            {
                Assert.AreEqual<int>(x, count);
                count += 2;
            });            
        }

        [TestMethod]
        public void IteratorFirstNElements()
        {
            var result = repository.GetAll().Take(10);
            var count = 0;
            result.ToList().ForEach(x=>Assert.AreEqual<int>(++count, x));
        }

        [TestMethod]
        public void IteratorFromSpecifiedIndexedElement()
        {
            var result = repository.GetAll().Skip(6);
            var count = 6;
            result.ToList().ForEach(x => Assert.AreEqual<int>(++count, x));            
        }

        [TestMethod]
        public void GetAvergage()
        {
            Assert.AreEqual<double>((1 + 15) / 2, repository.GetAll().Average());
        }
    }
}
