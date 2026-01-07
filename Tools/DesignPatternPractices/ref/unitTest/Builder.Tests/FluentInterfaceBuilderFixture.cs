using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Builder.Tests
{
    class Entry
    {
        public string A { get; set; }                   //  essential
        public int B { get; private set; }              //  essential
        public double C { get; set; }                   //  optional
        public float D { get; private set; }            //  optional
        public DateTime E { get; private set; }         //  optional
        public IList<string> F { get; set; }            //  optional

        /// <summary>
        /// 对外封闭Entry的构造过程
        /// </summary>
        /// <param name="builder"></param>
        private Entry(Builder builder)
        {
            A = builder.AValue;
            B = builder.BValue;
            C = builder.CValue;
            D = builder.DValue;
            E = builder.EValue;
            F = builder.FValue;
        }

        public class Builder
        {
            public string AValue { get; private set; }      //  essentail
            public int BValue { get; private set; }         //  essential

            public double CValue { get; private set; }      //  optional
            public float DValue { get; private set; }       //  optional
            public DateTime EValue { get; private set; }    //  optional
            public IList<string> FValue { get; private set;}//  optional

            public Builder(string a, int b)
            {
                AValue = a;
                BValue = b;

                //  C、D 默认为0, F 默认为 null
                EValue = DateTime.Now;

            }

            #region 连贯接口方法

            public Builder C(double c){ this.CValue = c; return this; }
            public Builder D(float d) { this.DValue = d; return this; }
            public Builder E(DateTime e) { this.EValue = e; return this; }
            public Builder F(IList<string> f) { this.FValue = f; return this; }

            #endregion

            /// <summary>
            /// 可以对Entry实例的构造过程进行管理
            /// 比如：重用既有的实例
            /// </summary>
            /// <returns></returns>
            public Entry BuildUp(){return new Entry(this);}
        }
    }

    [TestClass]
    public class FluentInterfaceBuilderFixture
    {
        [TestMethod]
        public void TestInnerFluentBuildUp()
        {
            var e1 = new Entry.Builder("a", 20)
                .C(30)
                .D(40)
                .F(new List<string> {"A", "B", "C"}).BuildUp();
            Output(e1);

            var e2 = new Entry.Builder("b", 18).BuildUp();
            Output(e2);
        }

        void Output(Entry entry)
        {
            Trace.WriteLine("\nEntry [" + entry.A + "]");
            Trace.WriteLine(entry.B.ToString());
            Trace.WriteLine(entry.C.ToString());
            Trace.WriteLine(entry.D.ToString());
            Trace.WriteLine(entry.E.ToString());
            Trace.WriteLine(entry.F == null ? "NULL" : string.Join(",", entry.F));            
        }
    }
}

namespace Raw
{
    class Entry
    {
        public string A { get; set; }
        public int B { get; private set; }
        public double C { get; set; }
        public float D { get; private set; }
        public DateTime E { get; private set; }
        public IList<string> F { get; set; }
    }
}
