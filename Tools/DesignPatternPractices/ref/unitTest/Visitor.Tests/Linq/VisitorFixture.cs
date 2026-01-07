using System.Linq;
using MarvellousWorks.PracticalPattern.Visitor.Delegating;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Visitor.Tests.Linq
{
    [TestClass]
    public class VisitorFixture
    {
        EmployeeCollection employees = new EmployeeCollection()
                                {
                                    new Employee{Name = "joe",Income = 25000,VacationDays = 14},
                                    new Manager{Name = "alice",Income = 22000,VacationDays = 14, Department ="sales" },
                                    new Employee{Name = "peter",Income = 15000,VacationDays = 7}
                                };

        void Validate()
        {
            var joe = employees[0];
            Assert.AreEqual<double>(25000 * 1.1, joe.Income);
            var peter = employees[2];
            Assert.AreEqual<int>(7 + 1, peter.VacationDays);

            var alice = employees[1];
            Assert.AreEqual<int>(14 + 2, alice.VacationDays);
            Assert.AreEqual<double>(22000 * 1.2, alice.Income);
        }

        [TestMethod]
        public void TestDelegateVisitor()
        {
            employees.Where(x => x.GetType() == typeof(Employee)).ToList().ForEach(x=>{ x.Income *= 1.1; x.VacationDays++;});
            employees.Where(x => x.GetType() == typeof(Manager)).ToList().ForEach(x => { x.Income *= 1.2; x.VacationDays+=2; });
            Validate();
        }
    }
}
