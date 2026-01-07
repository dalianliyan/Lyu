using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Visitor.Classic
{
    /// <summary>
    /// visitor 需要影响的 Element
    /// </summary>
    public interface IEmployee
    {
        /// <summary>
        /// 相关属性方法
        /// </summary>
        string Name { get; set;}
        double Income { get; set;}
        int VacationDays { get; set;}

        void Accept(IVisitor visitor);
    }

    /// <summary>
    /// 抽象Visitor接口
    /// </summary>
    public interface IVisitor
    {
        void VisitEmployee(Employee employee);
        void VisitManager(Manager manager);
    }   

    /// <summary>
    /// 一个具体的Element
    /// </summary>
    public class Employee : IEmployee
    {
        public Employee(string name, double income, int vacationDays)
        {
            Name = name;
            Income = income;
            VacationDays = vacationDays;
        }

        public string Name { get; set; }
        public double Income { get; set; }
        public int VacationDays { get; set; }

        /// <summary>
        /// 引入Visitor对自身的操作
        /// </summary>
        /// <param name="visitor"></param>
        public virtual void Accept(IVisitor visitor)
        {
            visitor.VisitEmployee(this);
        }
    }

    /// <summary>
    /// 另一个具体的Element
    /// </summary>
    public class Manager : Employee
    {
        public string Department { get; private set; }

        public Manager(string name, double income, int vacationDays, string department)
            : base(name, income, vacationDays)
        {
            Department = department;
        }

        public override void Accept(IVisitor visitor)
        {
            visitor.VisitManager(this);
        }
    }

    /// <summary>
    /// 为了便于对HR系统的对象进行批量处理增加的集合类型
    /// </summary>
    public class EmployeeCollection : List<IEmployee>
    {
        /// <summary>
        /// 组合起来的批量Accept操作
        /// </summary>
        /// <param name="visitor"></param>
        public virtual void Accept(IVisitor visitor)
        {
            if(visitor == null) 
                throw new ArgumentNullException("visitor");
            ForEach(x=>x.Accept(visitor));
        }
    }

    /// <summary>
    /// 具体的Visitor, 增加休假天数
    /// </summary>
    public class ExtraVacationVisitor : IVisitor
    {
        public void VisitEmployee(Employee employee)
        {
            employee.VacationDays += 1;
        }

        public void VisitManager(Manager manager)
        {
            manager.VacationDays += 2;
        }
    }
    /// <summary>
    /// 具体的Visitor, 加薪
    /// </summary>
    public class RaiseSalaryVisitor : IVisitor
    {
        public void VisitEmployee(Employee employee)
        {
            employee.Income *= 1.1;
        }

        public void VisitManager(Manager manager)
        {
            manager.Income *= 1.2;
        }
    }
}
