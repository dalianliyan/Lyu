namespace MarvellousWorks.PracticalPattern.Singleton.SimplestCounter
{
    public class Counter
    {
        Counter() { }
        public static readonly Counter Instance = new Counter();

        int value;
        public int Next { get { return ++value; } }
        public void Reset() { value = 0; }
    }
}
