using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Visitor.Classic;
namespace MarvellousWorks.PracticalPattern.Visitor.Tests.Classic
{
    [TestClass]
    public class VisitorFixture
    {
        [TestMethod]
        public void Test()
        {
            var employees = new EmployeeCollection()
                                {
                                    new Employee("joe", 25000, 14),
                                    new Manager("alice", 22000, 14, "sales"),
                                    new Employee("peter", 15000, 7)
                                };

            employees.Accept(new ExtraVacationVisitor());
            employees.Accept(new RaiseSalaryVisitor());

            var joe = employees[0];
            Assert.AreEqual<double>(25000 * 1.1, joe.Income);
            var peter = employees[2];
            Assert.AreEqual<int>(7 + 1, peter.VacationDays);

            var alice = employees[1];
            Assert.AreEqual<int>(14 + 2, alice.VacationDays);
            Assert.AreEqual<double>(22000 * 1.2, alice.Income);
        }
    }
}
