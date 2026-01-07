using System;
using System.Reflection;
namespace MarvellousWorks.PracticalPattern.Builder.Attributed
{
    /// <summary>
    /// 指导每个具体类型BuildPart过程目标方法和执行情况的属性
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple=false)]
    public sealed class BuildStepAttribute : Attribute
    {
        int sequence;
        int times;

        public BuildStepAttribute(int sequence, int times)
        {
            if((sequence <= 0) || (times <= 0 ))
                throw new ArgumentOutOfRangeException();
            this.sequence = sequence;
            this.times = times;
        }
        public BuildStepAttribute(int sequence) : this(sequence, 1) { }

        /// <summary>
        /// 执行的次序
        /// </summary>
        public int Sequence { get { return this.sequence; } }

        /// <summary>
        /// 执行的次数
        /// </summary>
        public int Times { get { return this.times; } }
    }
}
