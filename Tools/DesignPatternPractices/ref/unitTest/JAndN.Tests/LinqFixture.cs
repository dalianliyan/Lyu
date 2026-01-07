using System;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests.Properties;
namespace MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests
{
    [TestClass]
    public class LinqFixture
    {
        #region Attribute LINQ

        const string Mark = "technical";
        [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
        class CategoryAttribute : Attribute
        {
            public String Name { get; set; }
            public CategoryAttribute(string name){Name = name;}
        }

        [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
        class XAttribute : Attribute{}
        [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
        class YAttribute : Attribute{}

        class A
        {
            [Category(Mark)]
            [Category("M1")]
            [X]
            public int M1() {return 0; }

            [Category("M2")]
            [Y]
            [X]
            public int M2(int a) {return -1; }

            [Category(Mark)]
            public void M3(int a, int b) { }

            [Category(Mark)]
            [Category(Mark)]
            [Category("M4")]
            [X]
            public int M4(int a, int b) {return 0; }
        }

        [TestMethod]
        public void LinqToQueryReflectionInfo()
        {
            // outer query
            (from method in typeof(A).GetMethods()
             where
                // sub query
                 (from attribute in method.GetCustomAttributes(typeof(CategoryAttribute), false)
                      .AsEnumerable().Cast<CategoryAttribute>()
                  where string.Equals(Mark, attribute.Name)
                  select attribute)
                  .Count() > 0
                && (method.ReturnType == typeof(int))
             select method.Name)
                .ToList().ForEach(x => Trace.WriteLine(x));
        }

        #endregion

        #region XML LINQ

        const string TitleKeyWord = "笑看";
        const string ItemItem = "item";
        const string TitleItem = "title";
        const string DateItem = "pubDate";
        const string GuidItem = "guid";
        [TestMethod]
        public void QueryXmlFile()
        {
            (from item in XDocument.Parse(Resource.RssData).Descendants(ItemItem)
             where item.Element(TitleItem).Value.Contains(TitleKeyWord)
             orderby (DateTime)item.Element(DateItem) descending 
             select new
                        {
                            Name = item.Element(TitleItem).Value, 
                            Guid = item.Element(GuidItem).Value
                        }
             
             )
             .ToList().ForEach(x => Trace.WriteLine(x.Name + ":" + x.Guid));
        }

        #endregion

        #region LINQ Join clause

        const string PubXmlFileName = "authorPubHistory.xml";
        const string AuthorItem = "author";
        [TestMethod]
        public void JoinQuery()
        {
            (from item in XDocument.Parse(Resource.RssData).Descendants(ItemItem)
             join pubItem in XDocument.Parse(Resource.AuthorPubHistory).Descendants(ItemItem)
                 on item.Element(TitleItem).Value equals pubItem.Attribute(TitleItem).Value
             where item.Element(TitleItem).Value.Contains(TitleKeyWord)
             orderby (DateTime)item.Element(DateItem) descending
             select new
                        {
                            Author = pubItem.Attribute(AuthorItem).Value,
                            Name = item.Element(TitleItem).Value,
                            Guid = item.Element(GuidItem).Value
                        }

             )
             .ToList().ForEach(x => Trace.WriteLine(x.Author + "|" + x.Name + ":" + x.Guid));
        }

        #endregion
    }
}
