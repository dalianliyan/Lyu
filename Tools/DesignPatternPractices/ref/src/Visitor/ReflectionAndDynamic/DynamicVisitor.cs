namespace MarvellousWorks.PracticalPattern.Visitor.ReflectionAndDynamic
{
    public class DynamicVisitor : IVisitor
    {
        public void Visit(IEmployee employee)
        {
            Visit((dynamic)employee);
        }

        void Visit(Employee employee)
        {
            employee.VacationDays += 1;
            employee.Income *= 1.1;
        }

        void Visit(Manager employee)
        {
            employee.VacationDays += 2;
            employee.Income *= 1.2;
        }
    }
}
