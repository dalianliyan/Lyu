using System;

namespace MarvellousWorks.PracticalPattern.Interpreter.Classic
{
    /// <summary>
    /// 表示所有操作数
    /// </summary>
    public class Operand : IExpression
    {
        public int Number { get; set; }

        /// <summary>
        /// 根据操作符执行计算
        /// </summary>
        /// <param name="c"></param>
        public virtual void Evaluate(Context c)
        {
            switch (c.Operator)
            {
                case '\0': c.Value = Number; break;
                case '+': c.Value += Number; break;
                case '-': c.Value -= Number; break;
            }
        }
    }
}
