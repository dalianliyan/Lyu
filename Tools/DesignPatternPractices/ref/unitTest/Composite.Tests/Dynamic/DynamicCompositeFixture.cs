using System.Diagnostics;
using System.Xml.XPath;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Composite.Dynamic;
namespace MarvellousWorks.PracticalPattern.Composite.Tests.Dynamic
{
    [TestClass]
    public class DynamicCompositeFixture
    {
        dynamic corporate;

        [TestInitialize]
        public void Initialize()
        {
            corporate = new XComponent("corporate");
            corporate.president = "david";
            corporate.vicepresident = "vicepresident";
            corporate.sales = new XComponent("sales");
            corporate.market = new XComponent("market");
            corporate.branch = new XComponent("branch");
            corporate.sales.joe = "joe";
            corporate.sales.bob = "bob";
            corporate.market.judi = new XComponent("judi");
            corporate.branch.manager = "manager";
            corporate.branch.manager.title = "senior manager";
            corporate.branch.manager.gender = "female";
            corporate.branch.peter = "peter";
        }

        [TestMethod]
        public void TestDynamicXmlComponent()
        {
            string xml = (corporate).Element.ToString();
            Trace.WriteLine(xml);
        }
    }
}

