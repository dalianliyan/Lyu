using System;
using System.Linq;
using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Interpreter.Classic
{
    /// <summary>
    /// 解析器
    /// </summary>
    public class Calculator
    {
        public int Calculate(string expression)
        {
            if(string.IsNullOrWhiteSpace(expression))
                throw new ArgumentNullException("expression");
            var context = new Context();
            var tree = new List<IExpression>();
            // 词法和语法分析
            expression.ToCharArray().ToList().ForEach(
                x =>
                    {
                        if ((x == '+') || (x == '-'))
                            tree.Add(new Operator() {Op = x});
                        else
                            tree.Add(new Operand() { Number = (int)(x - 48) });
                    });
            // 遍历执行每个中间过程);)
            tree.ForEach(x => x.Evaluate(context));
            return context.Value;
        }
    }
}
