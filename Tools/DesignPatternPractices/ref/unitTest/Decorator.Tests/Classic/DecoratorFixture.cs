using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Decorator.Classic;
namespace MarvellousWorks.PracticalPattern.Decorator.Tests.Classic
{
    [TestClass]
    public class DecoratorFixture
    {
        [TestMethod]
        public void Test()
        {
            // 建立对象，并对其进行两次装饰
            IText text = new TextObject();
            text = new BoldDecorator(text);
            text = new ColorDecorator(text);
            Assert.AreEqual<string>("<color><b>hello</b></color>", text.Content);
            text = null;

            // 建立对象，只对其进行1次装饰
            text = new TextObject();
            text = new ColorDecorator(text);
            Assert.AreEqual<string>("<color>hello</color>", text.Content);

            // 通过装饰，撤销某些操作
            text = new BlockAllDecorator(text);
            Assert.IsTrue(string.IsNullOrEmpty(text.Content));
        }
    }
}
