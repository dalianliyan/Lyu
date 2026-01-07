using System;
namespace MarvellousWorks.PracticalPattern.Concept.Delegating
{
    #region 业务实体和事件监控器

    /// <summary>
    /// 业务实体
    /// </summary>
    public class Order 
    {
        public void Create() { EventMonitor.Added(this, null); }    // 新增
        public void ChangeDate() { EventMonitor.Modified(this, null); } // 修改
        public void ChangeOwner() { EventMonitor.Modified(this, null); } // 修改
        public void ChangeId() { }  // 不计入监控
    }

    /// <summary>
    /// 事件监控器
    /// </summary>
    public static class EventMonitor 
    {
        public static EventHandler<EventArgs> Modified;
        public static EventHandler<EventArgs> Added;

        static EventMonitor() 
        {
            Modified = OnModified;
            Added = OnAdded;
        }

        public static int ModifiedTimes { get;  private set; }
        public static int AddedTimes { get;  private set; }

        static void OnModified(object sender, EventArgs args) { ModifiedTimes++; }
        static void OnAdded(object sender, EventArgs args) { AddedTimes++; }
    }

    #endregion
}
