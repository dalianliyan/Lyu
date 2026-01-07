using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.JavaAndDotNet.Tests
{
    [TestClass]
    public class OperatorOverloadingFixture
    {
        public struct Complex
        {
            public int R { get; set; }
            public int I { get; set; }

            public static Complex operator +(Complex c1, Complex c2)
            {
                return new Complex()
                           {
                               R = c1.R + c2.R,
                               I = c1.I + c2.I
                           };
            }
            public static Complex operator +(Complex c1, int i)
            {
                return new Complex()
                           {
                               R = c1.R + i,
                               I = c1.I
                           };
            }

            public override bool Equals(object obj)
            {
                var target = (Complex) obj;
                return (this.R == target.R) && (this.I == target.I);
            }
        }

        [TestMethod]
        public void CalculateCompelxNumber()
        {
            var c1 = new Complex() {R = 1, I = 2};   // 1+2i
            var c2 = new Complex() {R = 3, I = -4};  // 3-4i
            Assert.AreEqual<Complex>(new Complex() {R = 10, I = -2},
                c1 + c2 + 6);    // 10-2i
        }
    }
}
