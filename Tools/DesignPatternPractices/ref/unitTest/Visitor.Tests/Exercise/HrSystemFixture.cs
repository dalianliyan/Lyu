using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Visitor.Tests.Exercise
{
    [TestClass]
    public class HrSystemFixture
    {
        #region Staff

        abstract class Person
        {
            public string Name { get; set; }
            public virtual double Income { get; set;}
            public virtual object Clone() {return MemberwiseClone();}
        }

        /// <summary>
        /// 外聘专家
        /// </summary>
        class Expert : Person
        {
            public Expert(double cost)
            {
                Cost = cost;
                Income = Cost;
            }
            public double Cost { get; set; }
        }

        abstract class Employee : Person { }

        class GeneralEmployee : Employee
        {
            public GeneralEmployee(double basicSalary, int workingYears)
            {
                BasicSalary = basicSalary;
                WorkingYears = workingYears;
            }
            public double BasicSalary { get; set; }
            public double ExtraSalary { get; set; }
            public int WorkingYears { get; set; }
            public int VacationDays { get; set; }

            public override double Income{get{return BasicSalary + ExtraSalary;}}
        }

        class Temporary : Employee
        {
            public double Wage { get; set; }
            public override double Income{get{return Wage;}}
        }

        class Mananger : GeneralEmployee
        {
            public Mananger(double basicSalary, int workingYears, string department)
                : base(basicSalary, workingYears)
            {
                Department = department;
            }
            public string Department { get; private set; }
        }

        #endregion

        #region HrRuntime

        class HrRuntime
        {
            
            
            /// <summary>
            /// 所有Element接受Visitor的统一入口
            /// </summary>
            /// <remarks>
            ///     Type    表示适用的人员类别
            ///     Action  表示Visitor扩展的功能
            /// </remarks>
            public IEnumerable<KeyValuePair<Type, Action<Person>>> Registry { get; set; }

            /// <summary>
            /// 所有Element
            /// </summary>
            public List<Person> Staff { get; set; }

            /// <summary>
            /// 调用各个适用的Visitor
            /// </summary>
            public void Visit()
            {
                if(Registry.Count() == 0) return;
                if(Staff.Count() == 0) return;

                //  动态生成传统Visit各个接口的过程
                //      VisitGeneralExpert()
                //      VisitGeneralTemporary()
                //      VisitGeneralGeneralEmployee()
                //      VisitGeneralManager()
                var visitorsByType = Registry.GroupBy(x => x.Key);

                //  动态执行Visit()过程
                Staff.ForEach(x => Registry.Where(a => a.Key == x.GetType()).ToList().ForEach(a => a.Value(x)));
            }

            public T FindClone<T>(string name)
                where T : Person
            {
                return (T)(Staff.FirstOrDefault(x=>string.Equals(x.Name, name)).Clone());
            }
        }


        HrRuntime runtime;
        const double MaxMistake = 0.1;

        #endregion

        [TestInitialize]
        public void Initialize()
        {
            Action<Person> workingYearsSalaryHandler =
                (x) =>
                    {
                        ((GeneralEmployee) x).ExtraSalary +=
                            ((GeneralEmployee) x).WorkingYears*50;
                    };
            Action<Person> calculatePensionAndHousingFundHandler =
                (x) =>
                    {
                        ((GeneralEmployee) x).ExtraSalary -= 
                            ((GeneralEmployee) x).Income*(0.15 + 0.05);
                    };
            Action<Person> calculateAnualLeave =
                (x) =>
                    {
                        ((GeneralEmployee) x).VacationDays = 
                            ((((GeneralEmployee) x).WorkingYears - 1 )/5 + 1)*5;
                    };

            #region Init Runtime

            runtime =
                new HrRuntime()
                {
                    Registry = new List<KeyValuePair<Type, Action<Person>>>()
                    {
                        //  给普通员工和经理分别加薪15%和10%
                        new KeyValuePair<Type, Action<Person>>(typeof (GeneralEmployee), x =>{((GeneralEmployee) x).BasicSalary *= 1.15;}),
                        new KeyValuePair<Type, Action<Person>>(typeof (Mananger), x =>{((Mananger)x).BasicSalary *= 1.1;}),

                        //  司龄工资：每年增加50元
                        new KeyValuePair<Type, Action<Person>>(typeof (GeneralEmployee), workingYearsSalaryHandler),
                        new KeyValuePair<Type, Action<Person>>(typeof (Mananger), workingYearsSalaryHandler),

                        //  保险和公积金：为了简化示例，养老保险和住房公积金分别按最后应发收入合计的5%和15%扣除
                        new KeyValuePair<Type, Action<Person>>(typeof (Temporary), x =>{((Temporary) x).Wage *=(1 -0.15 -0.05);}),
                        new KeyValuePair<Type, Action<Person>>(typeof(GeneralEmployee), calculatePensionAndHousingFundHandler),
                        new KeyValuePair<Type, Action<Person>>(typeof(Mananger), calculatePensionAndHousingFundHandler),

                        //  年休假：所有员工（普通员工、经理）按照司龄从每年5天起，每满5年年度增加5天逐级递增
                        new KeyValuePair<Type, Action<Person>>(typeof(GeneralEmployee), calculateAnualLeave),
                        new KeyValuePair<Type, Action<Person>>(typeof(Mananger), calculateAnualLeave),

                        //  岗位增休假：计划在现有休假基础上，普通员工每年增加3天，经理每年增加5天
                        new KeyValuePair<Type, Action<Person>>(typeof(GeneralEmployee), (x)=>((GeneralEmployee)x).VacationDays += 3),
                        new KeyValuePair<Type, Action<Person>>(typeof(Mananger), (x)=>((Mananger)x).VacationDays += 5),
                    },
                    #region Add Demo Staff
                    Staff = new List<Person>()
                    {
                        new Expert(2000){Name = "E1"},
                        new Expert(3000){Name = "E2"},
                        new Temporary(){Name = "T1", Wage = 1200},
                        new Temporary(){Name = "T2", Wage = 1100},
                        new GeneralEmployee(4000, 3){Name = "G1"},
                        new GeneralEmployee(6000, 10){Name = "G2"},
                        new GeneralEmployee(4000, 3){Name = "G3"},
                        new Mananger(11000, 2, "sales"){Name = "M1"},
                        new Mananger(15000, 12, "administration"){Name = "M2"}
                    }
                    #endregion
                };

            #endregion
        }


        /// <summary>
        /// 外聘专家
        /// </summary>
        [TestMethod]
        public void TestExpert()
        {
            var e1 = runtime.FindClone<Expert>("E1");
            var e2 = runtime.FindClone<Expert>("E2");

            runtime.Visit();

            //  外聘专家没有加薪、加休的内容
            Assert.AreEqual<double>(e1.Income, runtime.FindClone<Expert>("E1").Income);
            Assert.AreEqual<double>(e2.Income, runtime.FindClone<Expert>("E2").Income);
        }

        /// <summary>
        /// 临时员工
        /// </summary>
        [TestMethod]
        public void TestTemporary()
        {
            var t1 = runtime.FindClone<Temporary>("T1");
            var t2 = runtime.FindClone<Temporary>("T2");

            runtime.Visit();

            //  仅适用 =>  保险和公积金：为了简化示例，养老保险和住房公积金分别按最后应发收入合计的5%和15%扣除
            Assert.AreEqual(t1.Income * 0.8, runtime.FindClone<Temporary>("T1").Income, MaxMistake);
            Assert.AreEqual(t2.Income * 0.8, runtime.FindClone<Temporary>("T2").Income, MaxMistake);
        }

        /// <summary>
        /// 普通员工
        /// </summary>
        [TestMethod]
        public void TestGeneralEmployee()
        {
            var g1 = runtime.FindClone<GeneralEmployee>("G1");
            var g2 = runtime.FindClone<GeneralEmployee>("G2");
            var g3 = runtime.FindClone<GeneralEmployee>("G3");

            runtime.Visit();

            //  收入适用 =>  
            //      给普通员工和经理分别加薪15%和10%
            //      司龄工资：每年增加50元
            //      保险和公积金：为了简化示例，养老保险和住房公积金分别按最后应发收入合计的5%和15%扣除
            Func<GeneralEmployee, double> incomeHandler =
                (x) =>
                    {
                        var extraSalary = x.WorkingYears*50;
                        var basicSalary = x.BasicSalary;
                        basicSalary *= 1.15;
                        var pensionAndHousingFund = (extraSalary + basicSalary)*0.2;
                        return extraSalary + basicSalary - pensionAndHousingFund;
                    };
            Assert.AreEqual(incomeHandler(g1), runtime.FindClone<GeneralEmployee>("G1").Income );
            Assert.AreEqual(incomeHandler(g2), runtime.FindClone<GeneralEmployee>("G2").Income);
            Assert.AreEqual(incomeHandler(g3), runtime.FindClone<GeneralEmployee>("G3").Income);


            //  休假适用 =>  
            //      年休假：所有员工（普通员工、经理）按照司龄从每年5天起，每满5年年度增加5天逐级递增
            //      岗位增休假：计划在现有休假基础上，普通员工每年增加3天，经理每年增加5天
            Func<GeneralEmployee, int> vacationHandler =
            (x) =>
            {
                var workingYears = x.WorkingYears;
                var vacation = ((workingYears - 1 )/5 + 1)*5;
                return vacation + 3;
            };
            Assert.AreEqual(vacationHandler(g1), runtime.FindClone<GeneralEmployee>("G1").VacationDays);
            Assert.AreEqual(vacationHandler(g2), runtime.FindClone<GeneralEmployee>("G2").VacationDays);
            Assert.AreEqual(vacationHandler(g3), runtime.FindClone<GeneralEmployee>("G3").VacationDays);
        }

        /// <summary>
        /// 经理
        /// </summary>
        [TestMethod]
        public void TestManager()
        {
            var m1 = runtime.FindClone<Mananger>("M1");
            var m2 = runtime.FindClone<Mananger>("M2");

            runtime.Visit();

            //  收入适用 =>  
            //      给普通员工和经理分别加薪15%和10%
            //      司龄工资：每年增加50元
            //      保险和公积金：为了简化示例，养老保险和住房公积金分别按最后应发收入合计的5%和15%扣除
            Func<GeneralEmployee, double> incomeHandler =
                (x) =>
                {
                    var extraSalary = x.WorkingYears * 50;
                    var basicSalary = x.BasicSalary;
                    basicSalary *= 1.1;
                    var pensionAndHousingFund = (extraSalary + basicSalary) * 0.2;
                    return extraSalary + basicSalary - pensionAndHousingFund;
                };
            Assert.AreEqual(incomeHandler(m1), runtime.FindClone<Mananger>("M1").Income);
            Assert.AreEqual(incomeHandler(m2), runtime.FindClone<Mananger>("M2").Income);


            //  休假适用 =>  
            //      年休假：所有员工（普通员工、经理）按照司龄从每年5天起，每满5年年度增加5天逐级递增
            //      岗位增休假：计划在现有休假基础上，普通员工每年增加3天，经理每年增加5天
            Func<GeneralEmployee, int> vacationHandler =
            (x) =>
            {
                var workingYears = x.WorkingYears;
                var vacation = ((workingYears - 1) / 5 + 1) * 5;
                return vacation + 5;
            };
            Assert.AreEqual(vacationHandler(m1), runtime.FindClone<Mananger>("M1").VacationDays);
            Assert.AreEqual(vacationHandler(m2), runtime.FindClone<Mananger>("M2").VacationDays);
        }
    }
}
