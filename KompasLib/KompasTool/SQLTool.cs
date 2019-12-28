using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KompasLib.Tools
{
    public class SQLTool
    {
        public string ReturnValue(string sqlCmd, string columName, string connectionstr)
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectionstr))
            {
                string value = string.Empty;

                SqlCommand com = new SqlCommand
                {
                    Connection = sqlConnection
                };
                sqlConnection.Open();

                SqlDataReader reader = null;
                com.CommandText = sqlCmd;

                reader = com.ExecuteReader();
                if (reader.Read())
                    value = reader[columName].ToString();

                sqlConnection.Close();

                return value;
            }
        }

        public DataSet ComboDataSet(string sqlCmd, string tablename, string connectionstr)
        {
            DataSet ds = new DataSet();

            using (SqlConnection sqlConnection = new SqlConnection(connectionstr))
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlCmd, sqlConnection);
                sqlConnection.Open();
                da.Fill(ds, tablename);

                sqlConnection.Close();
            }
            return ds;
        }


    }
}
