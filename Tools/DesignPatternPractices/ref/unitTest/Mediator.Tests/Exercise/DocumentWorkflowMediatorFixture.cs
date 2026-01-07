using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Mediator.Tests.Exercise
{
    class Document
    {
        #region essential fields

        public string Subject { get; set; }
        public string Body { get; set; }

        #endregion

        #region optional fields

        public string Comment { get; set; }

        #endregion

        public override string ToString()
        {
            return string.Format("\n[{0}]\n------------------\n{1}\n({2})\n", Subject, Body, Comment);
        }
    }

    class DocumentEventArgs : EventArgs
    {
        Document document;
        public DocumentEventArgs(Document document)
        {
            this.document = document;
        }

        public Document Document{get{ return document;}}
    }

    abstract class Department
    {
        protected Document document = new Document();
        public Document Document
        {
            get{ return document;}
            protected set{ document = value;}
        }
    }

    namespace Scenario1
    {
        class A : Department
        {
            public event EventHandler<DocumentEventArgs> WriteDocumentFinishedHandler;

            public void Write(string subject, string body, string comment)
            {
                if (Document == null) throw new NullReferenceException("Document");
                Trace.WriteLine("A begin write");
                Document.Subject = subject;
                Document.Body = body;
                Document.Comment = comment;
                Trace.WriteLine("A write finished");
                Trace.WriteLine(Document);

                if(WriteDocumentFinishedHandler != null)
                    WriteDocumentFinishedHandler(this, new DocumentEventArgs(Document));
            }

            public void OnReviewFailed(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("A received doc review failed from " + sender.GetType().Name);
                Document = eventArgs.Document;
            }
        }

        class B : Department
        {
            public event EventHandler<DocumentEventArgs> ReviewDocumentSucceedHandler;
            public event EventHandler<DocumentEventArgs> ReviewDocumentFailedHandler;

            public void Review()
            {
                Trace.WriteLine("B begin review");
                var result =
                    (Document != null) &&
                    ! string.IsNullOrEmpty(Document.Subject) &&
                    ! string.IsNullOrEmpty(Document.Body);

                if (result)
                {
                    Trace.WriteLine("B review succeed");
                    if (ReviewDocumentSucceedHandler != null)
                        ReviewDocumentSucceedHandler(this, new DocumentEventArgs(Document));
                }
                else
                {
                    Trace.WriteLine("B review failed");
                    if (ReviewDocumentFailedHandler !=  null)
                        ReviewDocumentFailedHandler(this, new DocumentEventArgs(Document));
                }
            }

            public void OnReceiveFileToReview(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("B received doc from " + sender.GetType().Name + " to review");
                Document = eventArgs.Document;
                Review();
            }
        }

        class C : Department
        {
            public event EventHandler<DocumentEventArgs> DocumentPublishedHandler;
            public event EventHandler<DocumentEventArgs> DocumentArchivedHandler;

            public void Publish()
            {
                Trace.WriteLine("C published ");
                if(DocumentPublishedHandler != null)
                    DocumentPublishedHandler(this, new DocumentEventArgs(Document));
            }

            public void Archive()
            {
                Trace.WriteLine("C archived");
                if(DocumentArchivedHandler != null)
                    DocumentArchivedHandler(this, new DocumentEventArgs(Document));
            }

            public void OnReceiveFileToPublish(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("C received doc to publish from " + sender.GetType().Name);
                Document = eventArgs.Document;
                Publish();
            }

            public void OnReceiveFileToArchive(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("C received doc to archive from " + sender.GetType().Name);
                Document = eventArgs.Document;
                Archive();
            }
        }
    }

    namespace Scenario2
    {
        class A : Department
        {
            public event EventHandler<DocumentEventArgs> WriteDocumentFinishedHandler;

            public void Write(string subject, string body, string comment)
            {
                if (Document == null) throw new NullReferenceException("Document");
                Trace.WriteLine("A begin write");
                Document.Subject = subject;
                Document.Body = body;
                Document.Comment = comment;
                Trace.WriteLine("A write finished");
                Trace.WriteLine(Document);

                if(WriteDocumentFinishedHandler != null)
                    WriteDocumentFinishedHandler(this, new DocumentEventArgs(Document));
            }

            public void OnReviewFailed(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("A received doc review failed from " + sender.GetType().Name);
                Document = eventArgs.Document;
            }
        }

        class B : Department
        {
            public event EventHandler<DocumentEventArgs> ReviewDocumentSucceedHandler;
            public event EventHandler<DocumentEventArgs> ReviewDocumentFailedHandler;

            public void Review()
            {
                Trace.WriteLine("B begin review");
                var result =
                    (Document != null) &&
                    ! string.IsNullOrEmpty(Document.Subject) &&
                    ! string.IsNullOrEmpty(Document.Body);

                if (result)
                {
                    Trace.WriteLine("B review succeed");
                    if (ReviewDocumentSucceedHandler != null)
                        ReviewDocumentSucceedHandler(this, new DocumentEventArgs(Document));
                }
                else
                {
                    Trace.WriteLine("B review failed");
                    if (ReviewDocumentFailedHandler != null)
                        ReviewDocumentFailedHandler(this, new DocumentEventArgs(Document));
                }
            }

            public void OnReceiveFileToReview(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("B received doc from " + sender.GetType().Name + " to review");
                Document = eventArgs.Document;
                Review();
            }
        }

        class C : Department
        {
            public event EventHandler<DocumentEventArgs> DocumentPublishedHandler;

            public void Publish()
            {
                Trace.WriteLine("C published ");
                if(DocumentPublishedHandler != null)
                    DocumentPublishedHandler(this, new DocumentEventArgs(Document));
            }

            public void OnReceiveFileToPublish(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("C received doc to publish from " + sender.GetType().Name);
                Document = eventArgs.Document;
                Publish();
            }
        }

        class D : Department
        {
            public event EventHandler<DocumentEventArgs> DocumentArchivedHandler;

            public void Archive()
            {
                Trace.WriteLine("D archived");
                if(DocumentArchivedHandler != null)
                    DocumentArchivedHandler(this, new DocumentEventArgs(Document));
            }

            public void OnReceiveFileToArchive(object sender, DocumentEventArgs eventArgs)
            {
                Trace.WriteLine("C received doc to archive from " + sender.GetType().Name);
                Document = eventArgs.Document;
                Archive();
            }
        }
    }

    [TestClass]
    public class DocumentWorkflowMediatorFixture
    {
        Scenario1.A a1;
        Scenario1.B b1;
        Scenario1.C c1;
        Scenario2.A a2;
        Scenario2.B b2;
        Scenario2.C c2;
        Scenario2.D d2;

        EventMediatorBuilder scenario1Builder;
        EventMediatorBuilder scenario2Builder;

        /// <summary>
        /// 配置不同协调关系
        /// 实际项目中可以采用本章介绍的基于配置文件的定义方式
        /// </summary>
        [TestInitialize]
        public void Initialize()
        {
            a1 = new Scenario1.A();
            b1 = new Scenario1.B();
            c1 = new Scenario1.C();
            a2 = new Scenario2.A();
            b2 = new Scenario2.B();
            c2 = new Scenario2.C();
            d2 = new Scenario2.D();

            scenario1Builder = new EventMediatorBuilder()
                //  1.	A部门拟稿后将文件报B部门审核
                .AddConfig(typeof (Scenario1.A), typeof (Scenario1.B), "WriteDocumentFinishedHandler", "OnReceiveFileToReview")
                //  2.	B部门对于文件审核后，确认文件体例没有缺项后就通知C部门发布
                .AddConfig(typeof (Scenario1.B), typeof (Scenario1.C), "ReviewDocumentSucceedHandler", "OnReceiveFileToPublish")
                //  3.	如果B部门发现文件体例有缺项时，将文件返回给A部门重新修改
                .AddConfig(typeof (Scenario1.B), typeof (Scenario1.A), "ReviewDocumentFailedHandler", "OnReviewFailed")
                //  4.	C部门在接到B部门传来的文件时，先再发布，然后对其归档
                .AddConfig(typeof(Scenario1.C), typeof(Scenario1.C), "DocumentPublishedHandler", "OnReceiveFileToArchive");


            scenario2Builder = new EventMediatorBuilder()
                .AddConfig(typeof(Scenario2.A), typeof(Scenario2.B), "WriteDocumentFinishedHandler", "OnReceiveFileToReview")
                .AddConfig(typeof(Scenario2.A), typeof(Scenario2.D), "WriteDocumentFinishedHandler", "OnReceiveFileToArchive")
                .AddConfig(typeof(Scenario2.B), typeof(Scenario2.A), "ReviewDocumentFailedHandler", "OnReviewFailed")
                .AddConfig(typeof(Scenario2.B), typeof(Scenario2.D), "ReviewDocumentFailedHandler", "OnReceiveFileToArchive")
                .AddConfig(typeof(Scenario2.B), typeof(Scenario2.C), "ReviewDocumentSucceedHandler", "OnReceiveFileToPublish")
                .AddConfig(typeof(Scenario2.C), typeof(Scenario2.D), "DocumentPublishedHandler", "OnReceiveFileToArchive");
            
        }

        /// <summary>
        /// 测试手工定义事件中介者的交互关系
        /// </summary>
        [TestMethod]
        public void TestManualDefineEventMediatorInSucceedBranch()
        {
            //  用事件配置松散的响应关系
            a1.WriteDocumentFinishedHandler += b1.OnReceiveFileToReview;
            b1.ReviewDocumentFailedHandler += a1.OnReviewFailed;
            b1.ReviewDocumentSucceedHandler += c1.OnReceiveFileToPublish;
            b1.ReviewDocumentSucceedHandler += c1.OnReceiveFileToArchive;

            //  成功的路径
            a1.Write("a", "b", "c");

            //  验证修订后的内容曾经流转给了B
            Assert.AreEqual<string>("a", b1.Document.Subject);
            Assert.AreEqual<string>("b", b1.Document.Body);
            Assert.AreEqual<string>("c", b1.Document.Comment);

            //  验证修订后的内容也曾经流转给了C
            Assert.AreEqual<string>("a", c1.Document.Subject);
            Assert.AreEqual<string>("b", c1.Document.Body);
            Assert.AreEqual<string>("c", c1.Document.Comment);
        }

        /// <summary>
        /// 测试手工定义事件中介者的交互关系
        /// </summary>
        [TestMethod]
        public void TestManualDefineEventMediatorInFailedBranch()
        {
            //  用事件配置松散的响应关系
            a1.WriteDocumentFinishedHandler += b1.OnReceiveFileToReview;
            b1.ReviewDocumentFailedHandler += a1.OnReviewFailed;
            b1.ReviewDocumentSucceedHandler += c1.OnReceiveFileToPublish;
            b1.ReviewDocumentSucceedHandler += c1.OnReceiveFileToArchive;

            //  失败的路径
            a1.Write("a", "", "");

            //  验证确实文档曾经流转给了B
            Assert.AreEqual<string>("a", b1.Document.Subject);
            Assert.AreEqual<string>("", b1.Document.Body);
            Assert.AreEqual<string>("", b1.Document.Comment);

            //  验证文档并没有流转给C
            Assert.IsNull(c1.Document.Subject);
            Assert.IsNull(c1.Document.Body);
            Assert.IsNull(c1.Document.Comment);

            //  修正错误的内容，重新执行流程
            a1.Write("a", "b", "c");

            //  验证修订后的内容曾经流转给了B
            Assert.AreEqual<string>("a", b1.Document.Subject);
            Assert.AreEqual<string>("b", b1.Document.Body);
            Assert.AreEqual<string>("c", b1.Document.Comment);

            //  验证修订后的内容也曾经流转给了C
            Assert.AreEqual<string>("a", c1.Document.Subject);
            Assert.AreEqual<string>("b", c1.Document.Body);
            Assert.AreEqual<string>("c", c1.Document.Comment);
        }

        /// <summary>
        /// 测试通过Event Mediator Builder构造现有业务流程下的协作关系
        /// </summary>
        [TestMethod]
        public void TestScenario1()
        {
            //  通过Event Mediator以及配置信息建立三个部门Colleague间的协作关系
            //  所有协调关系统一剥离到作为Mediator的.NET事件机制上
            scenario1Builder.BuildAUpColleagues(a1, b1, c1);

            //  成功的路径
            Trace.WriteLine("Succeed path");
            a1.Write("a", "b", "c");

            //  失败的路径
            Trace.WriteLine("\n\nFailed path");
            a1.Write("a", "", "");

            //  修正错误的内容，重新执行流程
            Trace.WriteLine("Modified after review failed path");
            a1.Write("a", "b", "c");
        }

        /// <summary>
        /// 测试通过Event Mediator Builder构造管理层期望的未来业务流程下的协作关系
        /// </summary>
        [TestMethod]
        public void TestScenario2()
        {
            //  通过Event Mediator以及配置信息建立三个部门Colleague间的协作关系
            //  所有协调关系统一剥离到作为Mediator的.NET事件机制上
            scenario2Builder.BuildAUpColleagues(a2, b2, c2, d2);

            //  成功的路径
            Trace.WriteLine("Succeed path");
            a2.Write("a", "b", "c");

            //  失败的路径
            Trace.WriteLine("\n\nFailed path");
            a2.Write("a", "", "");

            //  修正错误的内容，重新执行流程
            Trace.WriteLine("Modified after review failed path");
            a2.Write("a", "b", "c");
        }
    }

    class EventMediatorBuilder
    {
        class ConfigItem
        {
            public Type SourceType { get; set; }
            public Type TargetType { get; set; }
            public string SourceEventName { get; set; }
            public string TargetHandlerMethodName { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null) throw new ArgumentNullException("obj");
                var target = (ConfigItem)obj;
                return
                    SourceType == target.SourceType &&
                    TargetType == target.TargetType &&
                    string.Equals(SourceEventName, target.SourceEventName) &&
                    string.Equals(TargetHandlerMethodName, target.TargetHandlerMethodName);
            }
        }

        IList<ConfigItem> config = new List<ConfigItem>();

        public EventMediatorBuilder AddConfig(Type sourceType, Type targetType, string sourceEventName, string targetHandlerMethodName)
        {
            if (sourceType == null) throw new ArgumentNullException("sourceType");
            if (targetType == null) throw new ArgumentNullException("targetType");
            if (string.IsNullOrEmpty(sourceEventName)) throw new ArgumentNullException("sourceEventName");
            if (string.IsNullOrEmpty(targetHandlerMethodName)) throw new ArgumentNullException("targetHandlerMethodName");

            if (sourceType.GetEvent(sourceEventName) == null) throw new NotSupportedException(sourceEventName);
            var item = new ConfigItem()
            {
                SourceType = sourceType,
                TargetType = targetType,
                SourceEventName = sourceEventName,
                TargetHandlerMethodName = targetHandlerMethodName
            };
            if (!config.Contains(item))
                config.Add(item);

            return this;
        }

        public EventMediatorBuilder BuildAUpColleagues(params object[] colleagues)
        {
            if (colleagues == null) throw new ArgumentNullException("colleagues");
            if (config.Count() == 0) return this;       //  没有通信关系配置项
            if (colleagues.Count() == 1) return this;    //  没有需要配置的关联对象组
            colleagues.ToList().ForEach(x => { if (x == null) throw new ArgumentNullException(); });

            ////  限制：不支持一类对象的某个实例同时向另一类对象多个实例的通知
            //if (colleagues.GroupBy(x => x.GetType()).Count() != colleagues.Count())
            //    throw new NotSupportedException();

            foreach (var item in config)
            {
                var sources = colleagues.Where(x => x.GetType() == item.SourceType);
                if ((sources == null) || (sources.Count() == 0))
                    continue;
                var targets = colleagues.Where(x => x.GetType() == item.TargetType);
                if ((targets == null) || (targets.Count() == 0))
                    continue;
                var eventInfo = item.SourceType.GetEvent(item.SourceEventName);
                if (eventInfo == null)
                    continue;
                var methodInfo = item.TargetType.GetMethod(item.TargetHandlerMethodName, BindingFlags.Public | BindingFlags.Instance);
                if (methodInfo == null)
                    continue;

                //  绑定事件响应关系
                foreach (var source in sources)
                    foreach (var target in targets)
                        eventInfo.AddEventHandler(source, Delegate.CreateDelegate(eventInfo.EventHandlerType, target, methodInfo));
            }

            return this;
        }
    }
}
