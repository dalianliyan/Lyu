using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Memento.Tests.Exercise
{
    [TestClass]
    public class ComprehensiveMementoFixture
    {
        interface IOriginator<T>
            where T : IState
        {
            /// <summary>
            /// 保存一个备忘
            /// </summary>
            /// <remarks>
            ///     默认方式——当前所有数据的全备忘
            /// </remarks>
            void SaveCheckPoint();

            /// <summary>
            /// 恢复一个备忘
            /// </summary>
            /// <remarks>
            ///     默认方式——恢复最近一次全备忘
            /// </remarks>
            void Undo();

            /// <summary>
            /// 保存一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            void SaveCheckPoint(string name);

            /// <summary>
            /// 恢复一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            void Undo(string name);

            /// <summary>
            /// 保存一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            /// <param name="version">备忘的版本</param>
            void SaveCheckPoint(string name, string version);

            /// <summary>
            /// 恢复一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            /// <param name="version">备忘的版本</param>
            void Undo(string name, string version);

            /// <summary>
            /// 保存一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            /// <param name="version">备忘的版本</param>
            /// <param name="subjectName">业务主题名称</param>
            void SaveCheckPoint(string name, string version, string subjectName);

            /// <summary>
            /// 恢复一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            /// <param name="version">备忘的版本</param>
            /// <param name="subjectName">业务主题名称</param>
            void Undo(string name, string version, string subjectName);

            /// <summary>
            /// 保存一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            /// <param name="version">备忘的版本</param>
            /// <param name="subjectName">业务主题名称</param>
            /// <param name="start">时间区间起点</param>
            /// <param name="end">时间区间终点</param>
            void SaveCheckPoint(string name, string version, string subjectName, DateTime start, DateTime end);

            /// <summary>
            /// 恢复一个备忘
            /// </summary>
            /// <param name="name">备忘的名称</param>
            /// <param name="version">备忘的版本</param>
            /// <param name="subjectName">业务主题名称</param>
            /// <param name="start">时间区间起点</param>
            /// <param name="end">时间区间终点</param>
            void Undo(string name, string version, string subjectName, DateTime start, DateTime end);
        }
    }
}
