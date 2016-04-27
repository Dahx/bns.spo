using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlServerCe;


namespace bns.data.sqlce
{
    class Database
    {
        string _connstring;

        public Database(string connectionstring)
        {
            _connstring = connectionstring;
        }


        public System.Data.DataTable GetItems(string qry)
        {
            using (SqlCeConnection conn = new SqlCeConnection(_connstring))
            {
                conn.Open();
                System.Data.DataSet ds = new System.Data.DataSet();
                SqlCeDataAdapter adpt = new SqlCeDataAdapter(qry, conn);
                adpt.Fill(ds);
                return ds.Tables[0];
            }
        }

        public string GetValue(string qry)
        {
            using (SqlCeConnection conn = new SqlCeConnection(_connstring))
            {
                conn.Open();
                SqlCeCommand cmd = new SqlCeCommand(qry);
                cmd.Connection = conn;
                object result = cmd.ExecuteScalar();
                return result.ToString();
            }
        }

    }
}
