using System;
using System.Text.RegularExpressions;
namespace MarvellousWorks.PracticalPattern.Interpreter.RegExpression
{
    /// <summary>
    /// 采用正则表达式方式表示的抽象接口
    /// </summary>
    public interface IRegExpression : IExpression
    {
        /// <summary>
        /// 是否匹配
        /// </summary>
        Func<string, bool> IsMatch { get; }
    }

    /// <summary>
    /// 采用正则表达式方式表示的抽象基类
    /// </summary>
    public abstract class RegExpressionBase : IRegExpression
    {
        protected Regex regex;
        protected RegExpressionBase(string expression)
        {
            regex = new Regex(expression, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            IsMatch = regex.IsMatch;
            Evaluate = EvaluateInternal;
        }

        public virtual Func<string, bool> IsMatch { get; private set; }
        public Action<Context> Evaluate{ get; private set;}


        /// <summary>
        /// 解析表达式
        /// </summary>
        /// <param name="context"></param>
        protected void EvaluateInternal(Context context)
        {
            if (context == null) throw new ArgumentNullException("context");
            switch (context.Operator)
            {
                case 'M':
                    EvaluateMatch(context);
                    break;
                case 'R':
                    EvaluateReplace(context);
                    break;
                default:
                    throw new ArgumentException();
            }
        }

        /// <summary>
        /// 通过Match方式解析表达式，可以被子类覆盖
        /// </summary>
        /// <param name="context"></param>
        protected virtual void EvaluateMatch(Context context)
        {
            context.Matches.Clear();
            var coll = regex.Matches(context.Content);
            if (coll.Count == 0) return;
            foreach (Match match in coll)
                context.Matches.Add(match.Value);
        }

        /// <summary>
        /// 通过Replace方法替换内容，可以被子类覆盖
        /// </summary>
        /// <param name="context"></param>
        protected virtual void EvaluateReplace(Context context)
        {
            context.Content = regex.Replace(context.Content, context.Replacement);
        }

    }
}
