using System.Diagnostics;
using System.Threading;
namespace MarvellousWorks.PracticalPattern.Concept.Delegating
{
    public class AsyncInvoker
    {
        public AsyncInvoker()
        {
            new Timer(new TimerCallback(OnTimerInterval), "slow", 2500, 2500);
            new Timer(new TimerCallback(OnTimerInterval), "fast", 2000, 2000);
            Trace.WriteLine("method");
        }

        static void OnTimerInterval(object state){Trace.WriteLine(state as string);}
    }
}
