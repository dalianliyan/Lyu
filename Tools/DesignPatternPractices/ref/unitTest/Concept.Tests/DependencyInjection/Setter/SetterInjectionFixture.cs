using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Concept.DependencyInjection;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.DependencyInjection.Setter
{

    [TestClass]
    public class SetterInjectionFixture
    {
        [TestMethod]
        public void Test()
        {
            var client = new Client() {TimeProvider = (new Assembler()).Create<ITimeProvider>()};
        }
    }
}
