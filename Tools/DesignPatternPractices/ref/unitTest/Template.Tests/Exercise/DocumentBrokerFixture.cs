using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.FactoryMethod;
namespace MarvellousWorks.PracticalPattern.Template.Tests.Exercise
{
    interface IDocument { }
    class D1 : IDocument { }
    class D2 : IDocument { }
    class D3 : IDocument { }

    class DocumentEventArgs<T> : EventArgs
        where T : IDocument
    {
        public T Document { get; set; }
    }

    interface IDocumentBroker<T>
        where T : IDocument
    {
        void Process(T document);
        event EventHandler<DocumentEventArgs<T>> DocumentSent;
    }

    /// <summary>
    /// 对外部看到的模板方法是一个方法void Process(T document)和一个事件event EventHandler<DocumentEventArgs<T>> DocumentSent
    /// 对于子类而言，内部模板是T ProcessInternal(T document)和 Send(T document)两个方法
    /// 而整个算法的模板定义在void Process(T document)中
    /// </summary>
    /// <typeparam name="T">文件实体类型</typeparam>
    abstract class DocumentBrokerBase<T> : IDocumentBroker<T>
        where T : IDocument
    {

        public event EventHandler<DocumentEventArgs<T>> DocumentSent;

        protected abstract void Send(T document);
        protected abstract T ProcessInternal(T document);

        public virtual void Process(T document)
        {
            if (document == null) throw new ArgumentNullException("document");

            #region 算法的模板

            Trace.WriteLine(string.Format("begin process document [{0}]", document.GetType().Name));
            document = ProcessInternal(document);

            Send(document);
            if (DocumentSent != null)
                DocumentSent(this, new DocumentEventArgs<T>() { Document = document });

            #endregion
        }
    }

    class D1DocumentBroker : DocumentBrokerBase<D1>
    {
        protected override D1 ProcessInternal(D1 document)
        {
            Trace.WriteLine("process D1 with COM+ Application");
            return document;
        }
        protected override void Send(D1 document) { Trace.WriteLine("send d1 to exchange server"); }
    }

    class D2DocumentBroker : DocumentBrokerBase<D2>
    {
        protected override D2 ProcessInternal(D2 document)
        {
            Trace.WriteLine("process D1 with .NET WF");
            return document;
        }
        protected override void Send(D2 document) { Trace.WriteLine("send d2 to .NET WCF Service"); }
    }

    class D3DocumentBroker : DocumentBrokerBase<D3>
    {
        protected override D3 ProcessInternal(D3 document)
        {
            Trace.WriteLine("process D1 with EJB3");
            return document;
        }
        protected override void Send(D3 document) { Trace.WriteLine("send d3 to WebLogic Server"); }
    }

    [TestClass]
    public class DocumentBrokerFixture
    {
        IFactory factory;

        [TestInitialize]
        public void Initialize()
        {
            factory = new Factory()
                .RegisterType<IDocumentBroker<D1>, D1DocumentBroker>()
                .RegisterType<IDocumentBroker<D2>, D2DocumentBroker>()
                .RegisterType<IDocumentBroker<D3>, D3DocumentBroker>();
        }

        [TestMethod]
        public void TestD1Broker()
        {
            var broker = factory.Create<IDocumentBroker<D1>>();
            broker.DocumentSent += OnDocumentSent<D1>;
            broker.Process(new D1());
        }

        [TestMethod]
        public void TestD2Broker()
        {
            var broker = factory.Create<IDocumentBroker<D2>>();
            broker.DocumentSent += OnDocumentSent<D2>;
            broker.Process(new D2());
        }

        [TestMethod]
        public void TestD3Broker()
        {
            var broker = factory.Create<IDocumentBroker<D3>>();
            broker.DocumentSent += OnDocumentSent<D3>;
            broker.DocumentSent += OnDocumentCrossPlatformSent<D3>;
            broker.Process(new D3());
        }

        /// <summary>
        /// 响应事件方式的模板“填充”内容
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sender"></param>
        /// <param name="eventArgs"></param>
        void OnDocumentSent<T>(object sender, DocumentEventArgs<T> eventArgs)
            where T : IDocument
        {
            Trace.WriteLine(string.Format("docuemnt [{0}] has been sent by {1}", 
                eventArgs.Document.GetType().Name,
                sender.GetType().Name));
        }

        /// <summary>
        /// 响应事件方式的模板“填充”内容
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sender"></param>
        /// <param name="eventArgs"></param>
        void OnDocumentCrossPlatformSent<T>(object sender, DocumentEventArgs<T> eventArgs)
            where T : IDocument
        {
            Trace.WriteLine("cross platform communication");
        }

    }
}

