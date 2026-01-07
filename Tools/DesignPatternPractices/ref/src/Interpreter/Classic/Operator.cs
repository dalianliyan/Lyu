namespace MarvellousWorks.PracticalPattern.Interpreter.Classic
{
    /// <summary>
    /// 表示所有操作符
    /// </summary>
    public class Operator : IExpression
    {
        public char Op { get; set; }
        public virtual void Evaluate(Context context) { context.Operator = Op; }
    }
}
