using System;

namespace MarvellousWorks.PracticalPattern.Interpreter.RegExpression
{
    /// <summary>
    /// 所有表达式的抽象接口
    /// </summary>
    public interface IExpression
    {
        /// <summary>
        /// 用Context保存中间结果
        /// </summary>
        Action<Context> Evaluate { get; }
    }
}
