using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Memento.Tests.Exercise
{
    [TestClass]
    public class MementoCoRFixture
    {
        /// <summary>
        /// 为了简化没有定义发起者接口，直接定义了发起者类型
        /// </summary>
        class Originator
        {
            public string State { get; set; }

            Memento m;

            public HandlerCoRBuilder<string> HandlerCoRBuilder { get; set; }

            /// <summary>
            /// 把备忘录定义为发起者的内部类型
            /// </summary>
            /// <typeparam name="T"></typeparam>
            class Memento
            {
                public string State { get; set; }
            }

            /// <summary>
            /// 把状态保存到备忘录
            /// </summary>
            public virtual void SaveCheckpoint()
            {
                Trace.WriteLine("\nOriginator.void SaveCheckpoint()\n-----------------");
                m = new Memento()
                        {
                            State = ProcessState(HandlerCoRBuilder.BuildUpProcessCoR(), this.State)
                        };
            }
            /// <summary>
            /// 从备忘录恢复之前的状态
            /// </summary>
            public virtual void Undo()
            {
                if (m == null) return;
                Trace.WriteLine("\nOriginator.void Undo()\n-----------------");
                State = ProcessState(HandlerCoRBuilder.BuildUpReverseCoR(), m.State);
            }

            string ProcessState(IEnumerable<Func<string, string >> processes, string state)
            {
                if(string.IsNullOrEmpty(state)) throw new ArgumentNullException("state");
                if(processes == null) return state;
                foreach (var process in processes)
                    state = process(state);
                return state;
            }
        }


        class HandlerCoRBuilder<T>
        {
            public HandlerCoRBuilder() { Handlers = new List<HandlerBase<T>>(); }
            public IList<HandlerBase<T>> Handlers { get; private set; }

            public IEnumerable<Func<T, T>> BuildUpProcessCoR()
            {
                if ((Handlers == null) || (Handlers.Count == 0)) return null;
                return Handlers.OrderBy(x => x.Sequence).Select(x => new Func<T, T>(x.Process));
            }

            public IEnumerable<Func<T, T>> BuildUpReverseCoR()
            {
                if ((Handlers == null) || (Handlers.Count == 0)) return null;
                return Handlers.OrderByDescending(x => x.Sequence).Select(x => new Func<T, T>(x.Reverse));
            }
        }

        /// <summary>
        /// 为了简化，假设State为string
        /// </summary>
        abstract class HandlerBase<T>
        {
            /// <summary>
            /// 操作执行的序列
            /// </summary>
            public int Sequence { get; set; }

            /// <summary>
            /// 正向操作内容
            /// </summary>
            /// <param name="target"></param>
            /// <returns></returns>
            public abstract T Process(T target);

            /// <summary>
            /// 逆向操作内容
            /// </summary>
            /// <param name="tareget"></param>
            /// <returns></returns>
            public abstract T Reverse(T tareget);
        }

        static Func<string, string, string> mementoProcessHandler =
            (x, y) =>
                {
                    var result = string.Format("<{0}>{1}</{2}>", y, x, y);
                    Trace.WriteLine(string.Format("{0} => {1}", x, result));
                    return result;
                };
        static Func<string, string, string> mementoReverseHandler =
            (x, y) =>
                {
                    var result = x.Replace("<" + y + ">", "").Replace("</" + y + ">", "");
                    Trace.WriteLine(string.Format("{0} => {1}", x, result));
                    return result;
                };


        class CompressHandler : HandlerBase<string>
        {
            public override string Process(string target) { return mementoProcessHandler(target, "c"); }
            public override string Reverse(string tareget) { return mementoReverseHandler(tareget, "c"); }
        }

        class EncryptHandler : HandlerBase<string>
        {
            public override string Process(string target) { return mementoProcessHandler(target, "e"); }
            public override string Reverse(string tareget) { return mementoReverseHandler(tareget, "e"); }
        }

        class SignHandler : HandlerBase<string>
        {
            public override string Process(string target) { return mementoProcessHandler(target, "s"); }
            public override string Reverse(string tareget) { return mementoReverseHandler(tareget, "s"); }
        }

        Originator originator;
        HandlerCoRBuilder<string> builder;

        [TestInitialize]
        public void Initialize()
        {
            builder = new HandlerCoRBuilder<string>();
            builder.Handlers.Add(new CompressHandler(){Sequence = 2});
            builder.Handlers.Add(new EncryptHandler() { Sequence = 3 });
            builder.Handlers.Add(new SignHandler() { Sequence = 1 });
            originator = new Originator()
                             {
                                 HandlerCoRBuilder = builder
                             };
        }
        
        [TestMethod]
        public void TestMementoAndCoRIntegration()
        {
            originator.State = "hello";
            originator.SaveCheckpoint();

            originator.State = "world";
            
            originator.Undo();
            Assert.AreEqual<string>("hello", originator.State);
        }
    }
}

