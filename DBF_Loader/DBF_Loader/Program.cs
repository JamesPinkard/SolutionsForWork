using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;


namespace DBF_Loader
{
    class Program
    {
        static void Main(string[] args)
        {
            string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=yourfilepath;Extended Properties=dBASE IV;User ID=Admin;Password=;";
            using (OleDbConnection con = new OleDbConnection(constr))
            {
                var sql = "select * from " + fileName;
                OleDbCommand cmd = new OleDbCommand(sql, con);
                con.Open();
                DataSet ds = new DataSet(); ;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
            }
        }
    }
}
