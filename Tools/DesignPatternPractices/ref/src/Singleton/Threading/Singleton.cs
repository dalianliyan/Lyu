using System;
namespace MarvellousWorks.PracticalPattern.Singleton.Threading
{
    public class Singleton
    {
        Singleton() { }

        [ThreadStatic]  // 说明每个Instance仅在当前线程内静态
        static Singleton instance;

        public static Singleton Instance
        {
            get
            {
                if (instance == null)
                    instance = new Singleton();
                return instance;
            }
        }
    }
}
