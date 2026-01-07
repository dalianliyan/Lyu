using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeFirst
{
    class Program
    {
        static void Main(string[] args)
        {
            Users u = new Users() { ID = 0 };            
            using (DBContextCodeFirst db = new DBContextCodeFirst())
            {
                db.Users.Add(u);
                db.SaveChanges();
            }
        }
    }
}
