using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Observer.Concept2;
namespace MarvellousWorks.PracticalPattern.Observer.Tests.Concept2
{
    [TestClass]
    public class ObserverFixture
    {
        [TestMethod]
        public void Test()
        {
            var x = new X();
            x.AddRange(new IUpdatableObject[] {new A(), new B(), new C()});
            x.Update(10);
            x.ForEach(t=>Assert.AreEqual<int>(10, t.Data));
        }
    }
}
