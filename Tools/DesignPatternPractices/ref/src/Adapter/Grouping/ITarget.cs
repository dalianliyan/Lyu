namespace MarvellousWorks.PracticalPattern.Adapter.Grouping
{
    /// <summary>
    /// Target
    /// </summary>
    public interface IDatabaseAdapter
    {
        /// <summary>
        /// Request()
        /// </summary>
        string ProviderName { get;}
    }

    /// <summary>
    /// Concrete Adapter
    /// </summary>
    public class OracleAdapter : IDatabaseAdapter
    {
        OracleDatabase adaptee = new OracleDatabase();
        public string ProviderName { get { return adaptee.GetDatabaseName(); } }
    }

    /// <summary>
    /// Concrete Adapter
    /// </summary>
    public class SqlServerAdapter : IDatabaseAdapter
    {
        SqlServerDatabase adaptee = new SqlServerDatabase();
        public string ProviderName { get { return adaptee.DbName; } }
    }
}
