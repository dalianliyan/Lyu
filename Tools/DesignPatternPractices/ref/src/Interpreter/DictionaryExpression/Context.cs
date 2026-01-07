namespace MarvellousWorks.PracticalPattern.Interpreter.DictionaryExpression
{
    /// <summary>
    /// 适用于字典信息的Context对象
    /// </summary>
    public class Context
    {
        public object Key{ get; set;}
        public object Value { get; set; }
        /// <summary>
        /// 'T'(to) 根据Value获得Key
        /// 'F'(from) 根据Key获得Value
        /// </summary>
        public char Operator { get; set; }
    }
}
