namespace MarvellousWorks.PracticalPattern.Interpreter.Classic
{
    /// <summary>
    /// 用于保存计算过程的中间结果以及当前执行的操作符
    /// </summary>
    public class Context
    {
        public int Value { get; set; }
        public char Operator { get; set; }
    }
}
