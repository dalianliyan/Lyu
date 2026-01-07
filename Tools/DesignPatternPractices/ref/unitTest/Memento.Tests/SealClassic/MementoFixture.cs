using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Memento.SealClassic;
namespace MarvellousWorks.PracticalPattern.Memento.Tests.SealClassic
{
    [TestClass]
    public class MementoFixture
    {
        [TestMethod]
        public void Test()
        {
            var originator = new Originator();
            // 保存发起者刚初始化后的状态
            originator.SaveCheckpoint();

            // 对发起者进行操作，验证状态的修改
            originator.IncreaseY();
            originator.DecreaseX();
            Assert.AreEqual<int>(-1, originator.Current.X);
            Assert.AreEqual<int>(1, originator.Current.Y);

            // 确认备忘录的恢复作用
            originator.Undo();
            Assert.AreEqual<int>(0, originator.Current.Y);
            Assert.AreEqual<int>(0, originator.Current.X);
        }
    }
}
