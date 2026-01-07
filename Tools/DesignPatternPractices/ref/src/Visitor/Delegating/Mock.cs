using System;
using System.Collections.Generic;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.Visitor.Delegating
{
    /// <summary>
    /// visitor 需要影响的 Element
    /// </summary>
    public abstract class EmployeeBase
    {
        /// <summary>
        /// 相关属性
        /// </summary>
        public virtual string Name { get; set;}
        public virtual double Income { get; set;}
        public virtual int VacationDays { get; set;}
    }

    /// <summary>
    /// 一个具体的Element， Visitable
    /// </summary>
    public class Employee : EmployeeBase { }
    public class Manager : Employee
    {
        public string Department { get; set; }
    }

    /// <summary>
    /// 为了便于对HR系统的对象进行批量处理增加的集合类型
    /// </summary>
    public class EmployeeCollection : List<EmployeeBase>{}
}
