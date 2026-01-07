using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Command.Tests.Exercise
{
    [TestClass]
    public class HierarchyCommandExerciseFixture
    {
        interface IEntity
        {
            int Id { get; set; }
            int Age { get; set; }
        }

        class Entity : IEntity
        {
            public int Id { get; set; }            
            public int Age { get; set; }
        }

        class EntityWithName : Entity
        {
            public string Name { get; set; }
        }

        /// <summary>
        /// 外部排序命令对象抽象定义
        /// </summary>
        Action<
            IEntity[],                      //  表示待排序数据
            Action<IEntity[], int, int>,    //  表示交换子命令对象抽象定义
            Func<IEntity, IEntity, bool>>   //  表示比较子命令对象抽象定义
            sortHandler;

        IEntity[] testData;

        /// <summary>
        /// 准备测试数据
        /// </summary>
        [TestInitialize]
        public void Initialize()
        {
            testData = new IEntity[]
            {
                new Entity(){Id = 1, Age = 20},
                new Entity(){Id=2, Age=18},
                new Entity(){Id=5, Age=12},
                new EntityWithName(){Id=4, Age=30},
                new Entity(){Id=10, Age=13}
            };
        }

        /// <summary>
        /// 验证对于ID属性方法的排序
        /// </summary>
        [TestMethod]
        public void TestSortById()
        {
            #region 定义外部排序命令对象

            //  冒泡算法排序
            sortHandler = (data, swapHandler, compareHandler) =>
                              {
                                  if ((data == null) || (data.Length == 0)) throw new ArgumentNullException("data");
                                  if (data.Length == 1) return;
                                  for (var i = 0; i < data.Length - 1; i++)
                                  {
                                      for (var j = i + 1; j < data.Length; j++)
                                          if (!compareHandler(data[i], data[j]))
                                              swapHandler(data, i, j);
                                  }
                              };

            #endregion

            #region 内部嵌套的命令对象

            Action<IEntity[], int, int>
                swap = (data, x, y) =>
                           {
                               var temp = data[x];
                               data[x] = data[y];
                               data[y] = temp;
                           };

            //  先定义为倒序排列
            Func<IEntity, IEntity, bool> compare = (x, y) => x.Id >= y.Id;

            #endregion

            //  验证倒序排列效果
            sortHandler(testData, swap, compare);
            Assert.IsNotNull(testData);
            for (var i = 0; i < testData.Length - 1; i++)
                Assert.IsTrue(testData[i].Id >= testData[i+1].Id);


            //  修改内部命令对象的定义，采用升序排列
            compare = (x, y) => x.Id <= y.Id;
            //  重新定义整体的命令对象
            sortHandler(testData, swap, compare);
            //  验证升序排列效果
            Assert.IsNotNull(testData);
            for (var i = 0; i < testData.Length - 1; i++)
                Assert.IsTrue(testData[i].Id <= testData[i + 1].Id);

        }
    }
}
