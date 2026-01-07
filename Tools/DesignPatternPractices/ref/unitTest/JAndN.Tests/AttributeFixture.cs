using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests
{
    [TestClass]
    public class AttributeFixture
    {
        [AttributeUsage(AttributeTargets.Class)]
        class EmployeeCategoryAttribute : Attribute
        {
            public String Name { get; set; }
            public EmployeeCategoryAttribute(string name)
            {
                Name = name;
            }
        }

        [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
        [EmployeeCategory("fulltime")]
        class FulltimeEmployeeAttribute : Attribute
        {
            public int BasicSalary { get; set; }
            public FulltimeEmployeeAttribute(int basicSalary) 
            {
                BasicSalary = basicSalary; 
            }
        }

        [FulltimeEmployee(20000)]
        class Manager{}

        [TestMethod]
        public void LoadAnnotationInfo()
        {
            Assert.AreEqual<string>("fulltime",((EmployeeCategoryAttribute) 
                typeof (FulltimeEmployeeAttribute).GetCustomAttributes(typeof (EmployeeCategoryAttribute), false)[0]).Name);
            Assert.AreEqual<int>(20000, ((FulltimeEmployeeAttribute) typeof (Manager).
                GetCustomAttributes(typeof (FulltimeEmployeeAttribute), false)[0]).BasicSalary);
        }
    }
}
