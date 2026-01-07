using System;
namespace MarvellousWorks.PracticalPattern.Concept.DependencyInjection.Example1
{
    class TimeProvider
    {
        public DateTime CurrentDate { get { return DateTime.Now; } }
    }

    public class Client
    {
        public int GetYear()
        {
            TimeProvider timeProvier = new TimeProvider();
            return timeProvier.CurrentDate.Year;
        }
    }
}
