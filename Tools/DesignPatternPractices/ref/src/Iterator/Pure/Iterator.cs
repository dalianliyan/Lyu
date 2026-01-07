using System;
using System.Linq;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Iterator.Pure
{
    public enum Color : int 
    {
        Red = 13, 
        Blue,
        Yellow
    }

    public class TreeNode<T>
    {
        public IList<TreeNode<T>> children { get; set; }
        public T Data { get; set; }

        #region Depth First Enumerator

        public IEnumerable<T> DepthFirstEnumerator
        {
            get
            {
                foreach (TreeNode<T> node in DepthFirstNodeEnumerator)
                    yield return node.Data;
            }
        }

        public IEnumerable<TreeNode<T>> DepthFirstNodeEnumerator
        {
            get
            {
                yield return this;
                if (children == null)
                    yield break;
                foreach (var child in children)
                {
                    var childEnum = child.DepthFirstNodeEnumerator.GetEnumerator();
                    while (childEnum.MoveNext())
                        yield return childEnum.Current;
                }
            }
        }
        #endregion
    }

    public class Repository
    {
        public int[] ArrayData { get; set; }
        public IEnumerable<KeyValuePair<int, int>> CheckList { get; set; }
        public TreeNode<int> Root { get; set; }

        /// <summary>
        /// 综合访问不同的迭代器
        /// </summary>
        /// <returns></returns>
        public IEnumerable<int> GetAll()
        {
            var result = new List<int>();
            if ((ArrayData != null) && (ArrayData.Length != 0))
                result.AddRange(ArrayData);
            if ((CheckList != null) || (CheckList.Count() != 0))
                result.AddRange(CheckList.Select(x => x.Value));
            if (Root != null)
                result.AddRange(Root.DepthFirstEnumerator);
            result.AddRange(Enum.GetValues(typeof (Color)).Cast<int>());
            return result.Count() == 0 ? null : result;
        }

        public int this[int index]
        {
            get { return GetAll().ElementAt(index); }
        }
    }
}
