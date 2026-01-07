using System.Linq;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Observer.Concept2
{
    /// <summary>
    /// 对类型A/B/C做抽象后的接口
    /// </summary>
    public interface IUpdatableObject
    {
        int Data { get;}
        void Update(int newData);
    }

    /// <summary>
    /// 具体的待更新类型
    /// </summary>
    public class A : IUpdatableObject
    {
        int data;
        public int Data{get { return this.data;}}
        public void Update(int newData) { this.data = newData; }
    }
    public class B : IUpdatableObject
    {
        int count;
        public int Data { get { return this.count; } }
        public void Update(int newData) { this.count = newData; }

    }
    public class C : IUpdatableObject
    {
        int n;
        public int Data { get { return this.n; } }
        public void Update(int newData) { this.n = newData; }
    }


    public class X : List<IUpdatableObject>
    {
        int data;
        public void Update(int newData)
        {
            this.data = newData;
            ForEach(x => x.Update(newData));
        }
    }
}
