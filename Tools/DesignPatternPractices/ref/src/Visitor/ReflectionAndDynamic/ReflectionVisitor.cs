using System;
namespace MarvellousWorks.PracticalPattern.Visitor.ReflectionAndDynamic
{
    public class ReflectionVisitor : IVisitor
    {
        const string Prefix = "Visit";

        /// <summary>
        /// 通过反射和预订好的方法命名规则动态执行
        /// </summary>
        /// <param name="employee"></param>
        public void Visit(IEmployee employee)
        {
            if (employee == null) throw new ArgumentNullException("employee");
            GetType().GetMethod(string.Format("{0}{1}", Prefix, employee.GetType().Name))
                .Invoke(this, new object[] { employee });
        }

        public void VisitEmployee(IEmployee employee)
        {
            employee.Income *= 1.1;
            employee.VacationDays += 1;
        }

        public void VisitManager(IEmployee employee)
        {
            var manager = (Manager)employee;
            manager.Income *= 1.2;
            manager.VacationDays += 2;
        }
    }
}
