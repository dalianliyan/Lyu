using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Iterator.Tests.Exercise
{
    [TestClass]
    public class AttributSubjectedIteratorFixture
    {
        [Serializable]
        [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
        class IterableAttribute : Attribute{}

        class SubjectAAttribute : IterableAttribute{}
        class SubjectBAttribute : IterableAttribute { }

        static class SubjectedIterator
        {
            /// <summary>
            /// 按照指定的主题——即Attribute遍历
            /// </summary>
            /// <typeparam name="TElement">集合元素类型</typeparam>
            /// <typeparam name="TAttribute">遍历所针对的Attribute类型</typeparam>
            /// <param name="data"></param>
            /// <returns></returns>
            public static IEnumerable<TElement> GetEnumerator<TElement, TAttribute>(IEnumerable<TElement> data)
                where TAttribute : IterableAttribute
            {
                if(data == null) return null;
                if (data.Count() == 0) return data;

                return data.Where(x =>
                                  x.GetType().GetCustomAttributes(typeof (TAttribute), false).Length > 0);
            }
        }


        [SubjectA]
        class A{}

        [SubjectB]
        class B{}

        [SubjectA]
        [SubjectB]
        class AB{}

        class O{}

        [TestMethod]
        public void TestIterateBySubject()
        {
            var data = new List<object> {new A(), new B(), new AB(), new A(), new O {}, new O {}, new AB()};

            var subjectA = SubjectedIterator.GetEnumerator<object, SubjectAAttribute>(data);
            Assert.AreEqual<int>(4, subjectA.Count());      //  两个A两个AB

            var subjectB = SubjectedIterator.GetEnumerator<object, SubjectBAttribute>(data);
            Assert.AreEqual<int>(3, subjectB.Count());      //  1个B两个AB
        }
    }
}
