using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Command.Tests.Exercise
{
    [TestClass]
    public class RewriteWikiJavaCommandSampleExerciseFixture
    {
        class Light
        {
            public void TurnOn(){Trace.WriteLine("the light is on");}
            public void TurnDown() { Trace.WriteLine("the light is off"); }    
        }

        class Switch
        {
            public Action UpCommand { private get; set; }
            public Action DownCommand { private get; set; }

            public void FlipUp() {UpCommand(); }
            public void FlipDown() { DownCommand(); }
        }

        /// <summary>
        /// 验证执行过程
        /// </summary>
        [TestMethod]
        public void Test()
        {
            var light = new Light();
            var s = new Switch()
                        {
                            UpCommand = light.TurnOn,
                            DownCommand = light.TurnDown
                        };
            s.FlipUp();
            s.FlipDown();
        }
    }
}
