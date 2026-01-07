using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.Exercise
{
    [TestClass]
    public class DependencyInjectionExerciseFixture
    {
        interface IFirst {}
        interface ISecond{}

        class Target
        {
            public IFirst First { get; set; }
            public ISecond Second { get; set; }
        }

        class First1 : IFirst{}
        class First2 : IFirst { }
        class First3 : IFirst { }

        class Second1 : ISecond{}
        class Second2 : ISecond{}

        [TestMethod]
        public void TestSetterInjection()
        {
            var target = new Target()
                             {
                                 First = new First1(),
                                 Second = new Second2()
                             };
            Assert.IsNotNull(target);
            Assert.IsInstanceOfType(target.First, typeof(First1));
            Assert.IsInstanceOfType(target.Second, typeof(Second2));
        }
    }
}
