using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Memento.Classic;
namespace MarvellousWorks.PracticalPattern.Memento.Tests.Classic
{
    [TestClass]
    public class MementoFixture
    {
        [TestMethod]
        public void Test()
        {
            var originator = new Originator();
            Assert.AreEqual<int>(0, originator.Current.Y);
            Assert.AreEqual<int>(0, originator.Current.X);
            // 保存发起者刚初始化后的状态
            var m1 = originator.Memento;

            // 对发起者进行操作，验证状态的修改
            originator.IncreaseY();
            originator.DecreaseX();
            Assert.AreEqual<int>(1, originator.Current.Y);
            Assert.AreEqual<int>(-1, originator.Current.X);

            // 确认备忘录的恢复作用
            originator.Memento = m1;
            Assert.AreEqual<int>(0, originator.Current.Y);
            Assert.AreEqual<int>(0, originator.Current.X);
        }
    }
}
