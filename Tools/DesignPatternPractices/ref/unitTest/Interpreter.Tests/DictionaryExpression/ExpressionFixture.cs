using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.Interpreter.DictionaryExpression;
namespace MarvellousWorks.PracticalPattern.Interpreter.Tests.DictionaryExpression
{
    [TestClass]
    public class ExpressionFixture
    {
        enum Color { Red, Green, Blue }
        
        IDictionary<string, string> data;
        SimpleDictionaryExpression expression;

        [TestInitialize]
        public void Initialize()
        {
            data = new Dictionary<string, string>()
                       {
                           {"R", "Red"},
                           {"G", "Green"},
                           {"B", "Blue"}
                       };
            expression = new SimpleDictionaryExpression();
        }

        [TestMethod]
        public void Test()
        {
            // 准备环境
            expression.Store = new EnumDictionaryStore<Color>();
            // 测试Enum从Key到Value的解析过程
            var context = new Context() { Key = Color.Red, Operator = 'F' };// from key to value
            expression.Evaluate(context);
            Assert.AreEqual<string>("Red", context.Value as string);

            // 测试Enum从Value到Key的解析过程
            context = new Context() { Value = "Blue", Operator = 'T' }; // from value to key
            expression.Evaluate(context);
            Assert.AreEqual<Color>(Color.Blue, (Color)(context.Key));

            // 测试对另一个DictionaryStore框架的解析能力
            expression.Store = new StringDictionaryStore() {Data = data};
            expression.Evaluate(context);
            Assert.AreEqual<string>("B", context.Key as string);
        }
    }
}
