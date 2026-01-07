using System;
using System.Collections.Generic;
using System.Linq;

namespace MarvellousWorks.PracticalPattern.Concept.Attributing
{
    public interface IAttributedBuilder
    {
        IList<string> Log { get; }      // log build steps
        void BuildPartA();
        void BuildPartB();
        void BuildPartC();
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class DirectorAttribute : Attribute, IComparable<DirectorAttribute>
    {
        public int Priority { get; set; }
        public string MethodName { get; set; }

        public DirectorAttribute(int priority, string methodName)
        {
            if (priority <= 0) throw new ArgumentOutOfRangeException("priority");
            if (string.IsNullOrEmpty(methodName))
                throw new ArgumentNullException("methodName");
            Priority = priority;
            MethodName = methodName;
        }

        public int CompareTo(DirectorAttribute attribute)
        {
            return attribute.Priority - this.Priority;
        }
    }

    public class Director
    {
        public void BuildUp(IAttributedBuilder builder)
        {
            ((DirectorAttribute[])
                builder.GetType().GetCustomAttributes(typeof(DirectorAttribute), false))
            .OrderByDescending(x => x.Priority)
            .ToList<DirectorAttribute>()
            .ForEach(x => InvokeBuildPartMethod(builder, x));
        }

        void InvokeBuildPartMethod(
            IAttributedBuilder builder, DirectorAttribute attribute)
        {
            switch (attribute.MethodName)
            {
                case "BuildPartA": builder.BuildPartA(); break;
                case "BuildPartB": builder.BuildPartB(); break;
                case "BuildPartC": builder.BuildPartC(); break;
            }
        }
    }
}
