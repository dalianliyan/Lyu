using System;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Observer.ObserverCollection.Simple
{
    /// <summary>
    /// 用于保存集合操作中操作条目信息的时间参数
    /// </summary>
    /// <typeparam name="TKey"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    public class DictionaryEventArgs<TKey, TValue> : EventArgs
    {
        public TKey Key{get;set;}
        public TValue Value { get; set; }
    }

    /// <summary>
    /// 具有操作事件的IDictionary<TKey, TValue>接口
    /// </summary>
    /// <typeparam name="TKey"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    public interface IObserverableDictionary<TKey, TValue> :
        IDictionary<TKey, TValue>
    {
        EventHandler<DictionaryEventArgs<TKey, TValue>> NewItemAdded { get; set;}
    }

    /// <summary>
    /// 一种比较简单的实现方式
    /// </summary>
    /// <typeparam name="TKey"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    public class ObserverableDictionary<TKey, TValue> : 
        Dictionary<TKey, TValue>, IObserverableDictionary<TKey, TValue>
    {
        public EventHandler<DictionaryEventArgs<TKey, TValue>> NewItemAdded{get; set;}

        /// <summary>
        /// 为既有操作增加事件
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public new void Add(TKey key, TValue value)
        {
            base.Add(key, value);
            if (NewItemAdded != null)
                NewItemAdded(this, new DictionaryEventArgs<TKey, TValue> {Key = key, Value = value});
        }
    }
}
