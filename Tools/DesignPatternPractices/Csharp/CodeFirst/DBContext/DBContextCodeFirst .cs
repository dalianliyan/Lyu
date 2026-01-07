using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace CodeFirst
{
    /// <summary>
    /// DBContext Class
    /// </summary>
    public class DBContextCodeFirst:DbContext
    {
        public DBContextCodeFirst() : base()
        {

        }

        #region DataTable

        public DbSet<Users> Users { get; set; }
        public DbSet<Movies> Movies { get; set; }

        #endregion //DataTable
    }
}
