using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Xml;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarvellousWorks.PracticalPattern.FactoryMethod;
namespace MarvellousWorks.PracticalPattern.Adapter.Tests.Exercise
{
    /// <summary>
    /// 队列访问接口
    /// </summary>
    interface IQueue
    {
        XmlDocument Peek();
        XmlDocument Dequeue();
    }

    /// <summary>
    /// 各在线业务数据实体的基类
    /// </summary>
    abstract class EntryBase { }


    /// <summary>
    /// 这里为了简化示例，将所有业务实体写如数据库前转换为ADO.NET自带的DataRow类型
    /// </summary>
    interface IEntryDataConverter
    {
        DataRow ToDataRow(EntryBase entry);
    }

    /// <summary>
    /// 数据库访问接口
    /// </summary>
    interface IDatabase
    {
        IEntryDataConverter DataConverter { get; set; }
        void Write(EntryBase entry);
        void Write(IEnumerable<EntryBase> entries);
    }

    [TestClass]
    public class MessageServiceFixture
    {

        #region 模拟用的队列产品接口及其适配器

        class Msmq
        {
            string current;

            /// <summary>
            /// 不兼容的接口，相当于Peek()
            /// </summary>
            public string Head
            {
                get
                {
                    Trace.WriteLine("Msmq.Head");
                    current = Guid.NewGuid().ToString();
                    return current;
                }
            }

            /// <summary>
            /// 不兼容的接口，相当于Dequeue()
            /// </summary>
            /// <returns></returns>
            public string GetHead()
            {
                Trace.WriteLine("Msmq.GetHead()");
                var result = current;
                current = string.Empty;
                return result;
            }
        }

        class MsmqAdapter : IQueue
        {
            Msmq queue = new Msmq();
            public XmlDocument Peek()
            {
                Trace.WriteLine("MsmqAdapter.Peek()");
                Trace.WriteLine(queue.Head);
                return null;
            }
            public XmlDocument Dequeue()
            {
                Trace.WriteLine("MsmqAdapter.Dequeue()");
                Trace.WriteLine(queue.GetHead());
                return null;
            }
        }

        class IbmMq : Queue<string>
        {
            string lastest;

            public new string Peek()
            {
                lastest = Guid.NewGuid().ToString();
                base.Enqueue(lastest);
                Trace.WriteLine("IbmMq.Peek()");
                return base.Peek();
            }

            public new string Dequeue()
            {
                Trace.WriteLine("IbmMq.Dequeue()");
                return base.Dequeue();
            }
        }

        class IbmMqAdapter : IQueue
        {
            IbmMq queue = new IbmMq();
            public XmlDocument Peek()
            {
                Trace.WriteLine("IbmMqAdapter.Peek()");
                Trace.WriteLine("Message : " + queue.Peek());
                return null;
            }
            public XmlDocument Dequeue()
            {
                Trace.WriteLine("IbmMqAdapter.Dequeue()");
                Trace.WriteLine("Message : " + queue.Dequeue());
                return null;
            }
        }

        #endregion

        #region 模拟用的数据库产品接口

        class OracleDB
        {
            public void WriteData(string[][] data)
            {
                Trace.WriteLine("OracleDb.WriteData(string[][] data)");
            }
        }

        class SqlServer
        {
            public void ExecuteNonSql(DataSet data)
            {
                Trace.WriteLine("SqlServer.ExecuteNonSql(DataSet data)");
            }

            public void InsertOneRecord(DataRow data)
            {
                Trace.WriteLine("SqlServer.InsertOneRecord(DataRow data)");
            }
        }

        class MySqlDatabase
        {
            public void ExecuteCommand(string sql)
            {
                Trace.WriteLine("MySqlDatabase.ExecuteCommand(string sql)");
            }
        }

        #endregion

        #region 模拟用的各类示例的业务实体

        class Order : EntryBase { }
        class DeliveryOrder : EntryBase { }
        class WarehouseEntry : EntryBase { }
        class InventoryList : EntryBase { }

        #endregion

        #region 模拟的数据适配接口以及面向个业务实体类型的实体数据适配类型

        interface IMessageConverter
        {
            IEnumerable<EntryBase> ConvertMessageToQuerable(XmlDocument message);
        }

        abstract class DatabaseAdapterBase : IDatabase
        {
            public IEntryDataConverter DataConverter { get; set; }

            public abstract void WriteData(DataRow row);
            public abstract void WriteData(IEnumerable<DataRow> rows);

            public virtual void Write(EntryBase entry)
            {
                if (DataConverter == null) throw new NullReferenceException("DataConverter");
                Trace.WriteLine("DatabaseAdapterBase.Write(EntryBase entry)");
                //if (entry != null)
                //    WriteData(DataConverter.ToDataRow(entry));
                WriteData(new DataTable().NewRow());
            }

            public virtual void Write(IEnumerable<EntryBase> entries)
            {
                if (DataConverter == null) throw new NullReferenceException("DataConverter");
                Trace.WriteLine("DatabaseAdapterBase.Write(IEnumerable<EntryBase> entries)");
                //if (entries != null)
                //{
                //    if (entries.Count() == 0) return;
                //    WriteData(entries.Select(x => DataConverter.ToDataRow(x)));
                //}
                WriteData(new List<DataRow>());
            }
        }

        class OralceAdapter : DatabaseAdapterBase
        {
            OracleDB db = new OracleDB();
            public override void WriteData(DataRow row)
            {
                Trace.WriteLine("OralceAdapter.WriteData(DataRow row)");
                //  将datarow变成string[0][], 然后提交OracleDb.WriteData(string[][] data)
                Trace.WriteLine("DataRow -> string[][]");
                db.WriteData(null);
            }

            public override void WriteData(IEnumerable<DataRow> rows)
            {
                Trace.WriteLine("OralceAdapter.WriteData(IEnumerable<DataRow> rows)");
                //  将IEnumerable<DataRow>变成string[][], 然后提交OracleDb.WriteData(string[][] data)
                Trace.WriteLine("IEnumerable<DataRow> -> string[][]");
                db.WriteData(null);
            }
        }

        class SqlAdapter : DatabaseAdapterBase
        {
            SqlServer db = new SqlServer();
            public override void WriteData(DataRow row)
            {
                Trace.WriteLine("SqlAdapter.WriteData(DataRow row)");
                //  直接将DataRow入库到SqlServer.InsertOneRecord(DataRow data)
                db.InsertOneRecord(null);
            }

            public override void WriteData(IEnumerable<DataRow> rows)
            {
                Trace.WriteLine("SqlAdapter.WriteData(IEnumerable<DataRow> rows)");
                //  将IEnumerable<DataRow>变成DataSet, 然后提交SqlServer.ExecuteNonSql(DataSet data)
                Trace.WriteLine("IEnumerable<DataRow> -> DataSet");
                db.ExecuteNonSql(null);
            }
        }

        class MySqlAdapter : DatabaseAdapterBase
        {
            MySqlDatabase db = new MySqlDatabase();
            public override void WriteData(DataRow row)
            {
                Trace.WriteLine("MySqlAdapter.WriteData(DataRow row)");
                //  直接将DataRow变成Insert SQL语句入库到MySqlDatabase.ExecuteCommand(string sql)
                Trace.WriteLine("DataRow -> SQL");
                db.ExecuteCommand(string.Empty);
            }

            public override void WriteData(IEnumerable<DataRow> rows)
            {
                Trace.WriteLine("MySqlAdapter.WriteData(IEnumerable<DataRow> rows)");
                //  将IEnumerable<DataRow>变成变成Insert SQL语句, 然后提交MySqlDatabase.ExecuteCommand(string sql)
                Trace.WriteLine("IEnumerable<DataRow> -> SQL");
                db.ExecuteCommand(string.Empty);
            }
        }

        /// <summary>
        /// 这里为了简化示例，定义了简单的转换类型，而且将只能分散的两个接口合二为一
        /// </summary>
        abstract class ConverterBase : IMessageConverter, IEntryDataConverter
        {
            public virtual IEnumerable<EntryBase> ConvertMessageToQuerable(XmlDocument message)
            {
                Trace.WriteLine("IMessageConverter  ConverterBase.ConvertMessageToQuerable(XmlDocument message)");
                return null;
            }

            public virtual DataRow ToDataRow(EntryBase entry)
            {
                Trace.WriteLine("IEntryDataConverter  ConverterBase.ToDataRow(EntryBase entry)");
                return null;
            }
        }

        class OrderConverter : ConverterBase { }
        class DeliveryOrderConverter : ConverterBase { }
        class WarehouseEntryConverter : ConverterBase { }
        class InventoryListConverter : ConverterBase { }

        #endregion

        #region 级联Queue Adapter和Db Adapter的连接器

        interface IMessageConnector
        {
            /// <summary>
            /// 处理的业务实体类型
            /// </summary>
            string EntryName { get; }

            /// <summary>
            /// 连接Queue Adapter和Db Adapter实现消息提取、转换和入库的过程
            /// </summary>
            void Process();
        }

        class MessageConnector : IMessageConnector
        {
            string entryName;
            IFactory factory;
            Route route;

            public MessageConnector(string entryName)
            {
                this.entryName = entryName;

                //  从运行环境获取配置信息
                factory = MessageServiceFixture.factory;
                route = MessageServiceFixture.orchestration.FirstOrDefault(x => string.Equals(x.EntryName, EntryName));
                if (route == null) throw new ApplicationException("route");
            }

            public string EntryName { get { return entryName; } }

            public virtual void Process()
            {
                Trace.WriteLine("MessageConnectorBase.Process");
                Trace.WriteLine(route);
                var queueAdapter = factory.Create<IQueue>(route.QueueName);
                //  为了示例简便，假设每个报文仅包括一批业务实体的情况
                //if (queueAdapter.Peek() != null)
                //{
                var message = queueAdapter.Peek();
                Trace.WriteLine("Get new message");
                var entries =
                    factory.Create<IMessageConverter>(EntryName)
                    .ConvertMessageToQuerable(queueAdapter.Dequeue());

                //  对业务实体进行检验和其他处理

                Trace.WriteLine("Save message to db");
                var dbAdpter = factory.Create<IDatabase>(route.DbName);
                //  注入
                dbAdpter.DataConverter = factory.Create<IEntryDataConverter>(EntryName);
                dbAdpter.Write(entries);
                //}
                Trace.WriteLine("\n\n");
            }
        }

        #endregion

        const string OracleItem = "ora";
        const string SqlServerItem = "sql";
        const string MySqlItem = "mysql";

        const string MsmqItem = "msmq";
        const string MqItem = "mq";

        const string OrderItem = "order";
        const string DeliveryOrderItem = "DeliveryOrder";
        const string WarehouseEntryItem = "WarehouseEntry";
        const string InventoryListItem = "InventoryList";


        /// <summary>
        /// 各类业务实体流转信息
        /// </summary>
        public class Route
        {
            /// <summary>
            /// 处理的业务实体名称
            /// </summary>
            public string EntryName { get; set; }

            /// <summary>
            /// 队列逻辑名称
            /// </summary>
            public string QueueName { get; set; }

            /// <summary>
            /// 数据库逻辑名称
            /// </summary>
            public string DbName { get; set; }

            public override string ToString()
            {
                return string.Format("{0} :\n---------------\nQueue\t\t{1}\nDatabase\t{2}", EntryName, QueueName, DbName);
            }
        }

        public static IFactory factory;

        /// <summary>
        /// 所有消息流转过程的登记清单
        /// </summary>
        public static IList<Route> orchestration;

        /// <summary>
        /// 加载各类配置信息
        /// </summary>
        /// <remarks>实际项目中这些配置一般登记在配置文件中</remarks>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            factory = new Factory()
                //  报文到业务实体的转换类型
                .RegisterType<IMessageConverter, OrderConverter>(OrderItem)
                .RegisterType<IMessageConverter, DeliveryOrderConverter>(DeliveryOrderItem)
                .RegisterType<IMessageConverter, WarehouseEntryConverter>(WarehouseEntryItem)
                .RegisterType<IMessageConverter, InventoryListConverter>(InventoryListItem)

                //  业务实体到DataRow的转换类型
                .RegisterType<IEntryDataConverter, OrderConverter>(OrderItem)
                .RegisterType<IEntryDataConverter, DeliveryOrderConverter>(DeliveryOrderItem)
                .RegisterType<IEntryDataConverter, WarehouseEntryConverter>(WarehouseEntryItem)
                .RegisterType<IEntryDataConverter, InventoryListConverter>(InventoryListItem)

                //  不同队列产品的适配器类型
                .RegisterType<IQueue, MsmqAdapter>(MsmqItem)
                .RegisterType<IQueue, IbmMqAdapter>(MqItem)

                //  不同数据库产品的适配器类型
                .RegisterType<IDatabase, OralceAdapter>(OracleItem)
                .RegisterType<IDatabase, SqlAdapter>(SqlServerItem)
                .RegisterType<IDatabase, MySqlAdapter>(MySqlItem);


            //  有订单（Order）、订单发货单（Order Delivery）、入库单（Warehouse entry）、库存清单（Inventory List）四种报文
            //	订单MSMQ接入，在线订单处理系统当前后台数据库为ORACLE
            //	入库单和库存清单通过IBM MQ接入，仓库后台处理系统数据库为SQL Server
            //	订单发货单通过IBM MQ接入，后台发货系统数据库为MySQL
            orchestration = new List<Route>()
                                {
                                    new Route()
                                        {
                                            EntryName = OrderItem,
                                            QueueName = MsmqItem,
                                            DbName = OracleItem
                                        },
                                    new Route()
                                        {
                                            EntryName = WarehouseEntryItem,
                                            QueueName = MqItem,
                                            DbName = SqlServerItem
                                        },
                                    new Route()
                                        {
                                            EntryName = InventoryListItem,
                                            QueueName = MqItem,
                                            DbName = SqlServerItem
                                        },
                                    new Route()
                                        {
                                            EntryName = DeliveryOrderItem,
                                            QueueName = MqItem,
                                            DbName = MySqlItem
                                        }
                                };
        }

        [TestMethod]
        public void TestDatabaseAdapter()
        {
            /// 启动服务端的所有处理通道
            List<IMessageConnector> channels =
                new List<IMessageConnector>()
                    {
                        new MessageConnector(OrderItem),
                        new MessageConnector(DeliveryOrderItem),
                        new MessageConnector(WarehouseEntryItem),
                        new MessageConnector(InventoryListItem)
                    };

            channels.ForEach(x=>x.Process());

        }
    }
}
