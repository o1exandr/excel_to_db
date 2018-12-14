using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;

namespace excel_to_db
{
    class UserProvider
    {
        private SqlConnection _con;

        public UserProvider(SqlConnection con)
        {
            _con = con;
        }
        
    }
}
