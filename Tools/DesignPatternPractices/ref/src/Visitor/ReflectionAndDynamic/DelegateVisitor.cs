using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Visitor.ReflectionAndDynamic
{
    public class DelegateVisitor : IVisitor
    {
        // delegate cache
        static readonly Dictionary<Type, Action<IEmployee>> actions;

        static DelegateVisitor()
        {
            // 有关委托的注册关系可以定义在配置文件中
            actions = new Dictionary<Type, Action<IEmployee>>()
                          {
                              {typeof (Employee), x => Visit((Employee) x)},
                              {typeof (Manager), x => Visit((Manager) x)}
                          };
        }

        public void Visit(IEmployee employee)
        {
            if(employee == null) throw new ArgumentNullException("employee");
            actions[employee.GetType()](employee);
        }

        static void Visit(Employee employee)
        {
            employee.VacationDays += 1;
            employee.Income *= 1.1;
        }

        static void Visit(Manager employee)
        {
            employee.VacationDays += 2;
            employee.Income *= 1.2;
        }
    }
}
