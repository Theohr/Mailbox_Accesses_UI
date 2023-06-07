using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.ComponentModel.Composition.Primitives;
using System.Windows.Forms;

namespace usersMailboxAccess
{
    internal class dbConn
    {
        //public static string SQL_CONNECT = @"Data Source=IOLNBTHNEW\SQLEXPRESS01;Initial Catalog=mailboxes;Integrated Security=True";
        public static string SQL_CONNECT = @"Data Source=server;Initial Catalog=mailboxes;Integrated Security=True";
        public static SqlConnection conn = new SqlConnection();
        public static SqlCommand sqlCmd = new SqlCommand();
        public static SqlDataReader sqlRdr;

        public static DataTable retrieveDB(string sqlCommand)
        {
            DataTable usersDataSet = new DataTable();

            SqlConnection cn = new SqlConnection(SQL_CONNECT);
            SqlDataAdapter daFeatures = new SqlDataAdapter(sqlCommand, cn);

            daFeatures.SelectCommand.CommandTimeout = 0;

            daFeatures.Fill(usersDataSet);

            return usersDataSet;
        }
    }
}
