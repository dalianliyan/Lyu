using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeFirst
{
    /// <summary>
    /// Model Class(POCO, Plain Old CLR Object)
    /// </summary>
    public class Users
    {
        /// <summary>
        /// int ID for PK, 必须有一个字段为主键，不然会报错
        /// 
        /// </summary>         
        public int ID { get; set; }

        public string UserName { get; set; }
    }
}
