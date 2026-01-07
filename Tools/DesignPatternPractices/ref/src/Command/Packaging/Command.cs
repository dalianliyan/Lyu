using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
namespace MarvellousWorks.PracticalPattern.Command.Packaging
{
    public enum DbOperation
    {
        Insert,
        Update,
        Delete
    }

    public class DbOperEventArgs : EventArgs
    {
        public int ID { get; set; }
        public DbOperation Operation { get; set; }
        public string NewValue { get; set; }

        public override string ToString()
        {
            return Operation == DbOperation.Update
                       ? string.Format("{0} {1} {2}", ID, Operation, NewValue)
                       : string.Format("{0} {1}", ID, Operation);
        }
    }

    /// <summary>
    /// Data transfer object for data operation
    /// </summary>
    public class DbOperDT
    {
        static int counter;
        List<DbOperEventArgs> Log = new List<DbOperEventArgs>();
        EventHandler<DbOperEventArgs> Changed { get; set; }

        public DbOperDT Update(int recordId, string newValue)
        {
            Log.Add(new DbOperEventArgs() { ID = recordId, Operation = DbOperation.Update, NewValue = newValue });
            return this;
        }

        public int Insert(string value)
        {
            Log.Add(new DbOperEventArgs() { ID = ++counter, Operation = DbOperation.Insert, NewValue = value });
            return counter;
        }

        public void Delete(int recordId)
        {
            Log.Add(new DbOperEventArgs() { ID = recordId, Operation = DbOperation.Delete });
        }

        public void Profile()
        {
            Log.ForEach(x=>Trace.WriteLine(x));
        }

        public void Merge()
        {
            if (Log.Count() <= 1) return;

            #region merge
            EraseInsertThenDelete();
            EraseUpdateThenDelete();
            MergeUpdateFinalResult();
            #endregion
        }

        /// <summary>
        /// 如果增加后在一个周期内删除的，直接擦除
        /// </summary>
        void EraseInsertThenDelete()
        {
            var list =
                from item in Log
                where item.Operation == DbOperation.Insert || item.Operation == DbOperation.Delete
                group item by item.ID
                    into g
                    where g.Count() == 2 // 1个Insert，一个Delete
                    select g.Key;
            if (list.Count() > 0)
                Log.RemoveAll(x => list.Contains(x.ID));
        }

        /// <summary>
        /// 如果多次更新后删除，则只保留删除操作，擦出更新操作
        /// </summary>
        void EraseUpdateThenDelete()
        {
            var list =
                (from item in Log
                 where item.Operation == DbOperation.Update || item.Operation == DbOperation.Delete
                 group item by item.ID
                     into g
                     where g.Count(x => x.Operation == DbOperation.Delete) > 0 && g.Count() >= 2
                     select g.Key).Distinct();
            if (list.Count() > 0)
                Log.RemoveAll(x => x.Operation == DbOperation.Update && list.Contains(x.ID));
        }

        /// <summary>
        /// 多次更新只保存终值
        /// </summary>
        void MergeUpdateFinalResult()
        {
            var list =
                from item in Log
                where item.Operation == DbOperation.Update
                group item by item.ID
                    into g
                    where g.Count() > 1
                    select g.Last();
            var ids = list.Select(x => x.ID).Distinct();
            var updates = list.ToArray();
            // 先清理掉所有更新记录
            Log.RemoveAll(x => ids.Contains(x.ID) && x.Operation != DbOperation.Insert);
            //  然后把最新更新的记录补充进去
            Log.AddRange(updates);
        }
    }
}
