using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Concept.Tests.Exercise
{
    class Notation
    {
        public Notation(){Data = string.Empty;}
        public Notation(string data) {Data = data; }
        public string Data { get; private set; }
    }

    /// <summary>
    /// n元素
    /// </summary>
    class Item : Notation
    {
        public Item():base(){}
        public Item(string data) : base(data){}
    }

    /// <summary>
    /// col 元素
    /// </summary>
    class Column : Notation
    {
        public Column():base(){}
        public Column(string data) : base(data) { }
    }

    /// <summary>
    /// line 元素 
    /// </summary>
    class Line
    {
        FluentCollection<Item, Line> items;
        Body body;

        public Line(Body body)
        {
            if(body == null) throw new ArgumentNullException("body");
            this.body = body;
            items = new FluentCollection<Item, Line>(this)
                        {
                            GetInstance = () => { return new Item(); }
                        };
        }

        /// <summary>
        /// 父节点
        /// </summary>
        public Body Body { get { return body; } }
        
        public FluentCollection<Item, Line> Items { get { return items; } }

        public Line NewLine{get{return body.NewLine;}}
    }


    /// <summary>
    /// body 元素
    /// </summary>
    class Body : WithTableObject
    {
        List<Line> lines = new List<Line>();
        public Body(Table table) : base(table){}

        public Line NewLine
        {
            get
            {
                var line = new Line(this);
                lines.Add(line);
                return line;
            }
        }

        public List<Line> Lines { get { return lines;}}
    } 

    /// <summary>
    /// head 元素
    /// </summary>
    class Head : WithTableObject
    {
        FluentCollection<Column, Head> columns;

        public Head(Table table) : base(table)
        {
            columns = new FluentCollection<Column, Head>(this)
                          {
                              GetInstance = () => { return new Column(); }
                          };
        }
        
        public FluentCollection<Column, Head> Columns { get { return columns; } }
    }

    
    class Table
    {
        string name;
        Body body;
        Head head;

        public Table()
        {
            body = new Body(this);
            head = new Head(this);
        }

        public Table Name(string name)
        {
            if(string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
            this.name = name;
            return this;
        }

        public override string ToString(){return name;}

        public Body Body{get{ return body;}}
        public Head Head{get{ return head;}}
    }

    /// <summary>
    /// 修改后具有Fluent特征的集合类型
    /// </summary>
    /// <typeparam name="T">集合元素类型</typeparam>
    /// <typeparam name="TParent">父节点类型</typeparam>
    class FluentCollection<TElement, TParent>
        where TElement : class 
        where TParent : class 
    {
        protected List<TElement> list = new List<TElement>();
        TParent parent;

        public FluentCollection(TParent parent)
        {
            if(parent == null) throw new ArgumentNullException("parent");
            this.parent = parent;
        }

        /// <summary>
        /// 返回父节点
        /// </summary>
        public TParent Parent{get{ return parent;}}

        /// <summary>
        /// 如何获得一个TElement类型实例的委托
        /// </summary>
        public Func<TElement> GetInstance { get; set; }

        /// <summary>
        /// 具有fluent特征的追加操作
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public FluentCollection<TElement, TParent> Add(TElement t)
        {
            list.Add(t);
            return this;
        }

        /// <summary>
        /// 具有fluent特征的空置操作
        /// </summary>
        /// <returns></returns>
        public FluentCollection<TElement, TParent> Skip
        {
            get
            {
                list.Add(GetInstance());
                return this;
            }
        }

        /// <summary>
        /// 执行LINQ的foreach操作
        /// </summary>
        /// <param name="action"></param>
        public void ForEach(Action<TElement> action)
        {
            list.ForEach(action);
        }
    }

    /// <summary>
    /// 父节点为table的元素
    /// </summary>
    class WithTableObject
    {
        Table table;    //  父节点
        public WithTableObject(Table table)
        {
            if(table == null) throw new ArgumentNullException("table");
            this.table = table;
        }

        /// <summary>
        /// 指向父节点——table
        /// </summary>
        public Table Parent{get{ return table;}}
    }

    class TableWriter
    {
        public void Output(Table table)
        {
            if (table == null) return;

            Trace.WriteLine("<table>");
            Trace.Indent();
            Trace.WriteLine(string.Format("<name>{0}</name>", table));
            Trace.WriteLine("<head>");
            if (table.Head.Columns != null)
            {
                Trace.Indent();
                table.Head.Columns.ForEach(x => Trace.Write(string.Format("<col>{0}</col>", x.Data)));
                Trace.Unindent();
                Trace.WriteLine("");
            }
            Trace.WriteLine("</head>");
            Trace.WriteLine("<body>");
            if (table.Body.Lines != null)
            {
                Trace.Indent();
                Trace.WriteLine("<line>");
                table.Body.Lines.ForEach(
                    x=>
                        {
                            Trace.WriteLine("<item>");
                            Trace.Indent();
                            if (x.Items != null)
                                x.Items.ForEach(y => Trace.Write(string.Format("<n>{0}</n>", y.Data)));
                            Trace.WriteLine("");
                            Trace.Unindent();
                            Trace.WriteLine("</item>");
                        }
                    );
                Trace.Unindent();
                Trace.WriteLine("</line>");
            }
            Trace.WriteLine("</body>");
            Trace.Unindent();
            Trace.WriteLine("</table>");

        }
    }

    [TestClass]
    public class FluentInterfaceFixture
    {
        TableWriter writer;

        [TestInitialize]  
        public void Initialize()
        {
            writer = new TableWriter();
        }

        [TestMethod]
        public void TestFullFillTable()
        {
            writer.Output(
                new Table()
                    .Name("full fill")
                    .Head
                        .Columns
                            .Add(new Column("first"))
                            .Add(new Column("second"))
                            .Add(new Column("thrid"))
                        .Parent
                    .Parent
                    .Body
                        .NewLine.Items.Add(new Item("11")).Add(new Item("12")).Add(new Item("13")).Parent
                        .NewLine.Items.Add(new Item("21")).Add(new Item("22")).Add(new Item("23")).Parent
                    .Body
                .Parent
                );
        }

        [TestMethod]
        public void TestSkipColumnTable()
        {
            writer.Output(
                new Table()
                    .Name("skip columns")
                    .Head
                        .Columns
                            .Add(new Column("first"))
                            .Skip
                            .Add(new Column("thrid"))
                        .Parent
                    .Parent
                    .Body
                        .NewLine.Items.Add(new Item("11")).Add(new Item("12")).Add(new Item("13")).Parent
                        .NewLine.Items.Add(new Item("21")).Add(new Item("22")).Add(new Item("23")).Parent
                    .Body
                .Parent
                );
        }

        [TestMethod]
        public void TestSkiItemsTable()
        {
            writer.Output(
                new Table()
                    .Name("skip items")
                    .Head
                        .Columns
                            .Add(new Column("first"))
                            .Add(new Column("second"))
                            .Add(new Column("thrid"))
                        .Parent
                    .Parent
                    .Body
                        .NewLine.Items.Add(new Item("11")).Skip.Add(new Item("13")).Parent
                        .NewLine.Items.Add(new Item("21")).Add(new Item("22")).Skip.Parent
                    .Body
                .Parent
                );
        }


        [TestMethod]
        public void TestSkipColumnsAndItemsTable()
        {
            writer.Output(
                new Table()
                    .Name("skip columns and items")
                    .Head
                        .Columns
                            .Add(new Column("first"))
                            .Skip
                            .Add(new Column("thrid"))
                        .Parent
                    .Parent
                    .Body
                        .NewLine.Items.Add(new Item("11")).Skip.Add(new Item("13")).Parent
                        .NewLine.Items.Add(new Item("21")).Add(new Item("22")).Skip.Parent
                    .Body
                .Parent
                );
        }
    }



}
