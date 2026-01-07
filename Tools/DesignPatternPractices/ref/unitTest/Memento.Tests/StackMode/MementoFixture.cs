using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Memento.StackMode;
namespace MarvellousWorks.PracticalPattern.Memento.Tests.StackMode
{
    [TestClass]
    public class MementoFixture
    {
        [TestMethod]
        public void Test()
        {
            var originator = new Originator();
            // 保存发起者刚初始化后的状态
            originator.SaveCheckpoint();    // x = 0, y = 0 入栈

            // 对发起者进行修改 1
            originator.IncreaseY();
            originator.DecreaseX();
            originator.SaveCheckpoint();    // x = -1, y = 1 入栈

            // 对发起者进行修改 2
            originator.UpdateX(2);
            originator.SaveCheckpoint();    // x = 2, y = 1 入栈

            // 对发起者进行修改 3
            originator.UpdateX(3);
            originator.IncreaseY();         // x = 3, y = 2
            

            // 确认Undo前的状态
            Assert.AreEqual<int>(3, originator.Current.X);
            Assert.AreEqual<int>(2, originator.Current.Y);
            

            // 确认备忘录的栈式恢复作用——恢复到修改
            originator.Undo();  //  出栈
            Assert.AreEqual<int>(2, originator.Current.X);
            Assert.AreEqual<int>(1, originator.Current.Y);

            // 确认备忘录的栈式恢复作用——恢复到修改1
            originator.Undo();  //  出栈
            Assert.AreEqual<int>(-1, originator.Current.X);
            Assert.AreEqual<int>(1, originator.Current.Y);
        }
    }
}
