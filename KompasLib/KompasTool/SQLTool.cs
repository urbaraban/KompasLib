using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

using System.Windows;

namespace KompasLib.Tools
{
    public static class SQLTool
    {
        public static string ReturnValue(string sqlCmd, string columName, string connectionstr)
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

        public static int GetLastAddID(string columName, string tableName, string connectionstr)
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectionstr))
            {
                int value = -1;

                SqlCommand com = new SqlCommand
                {
                    Connection = sqlConnection
                };
                sqlConnection.Open();
                com.CommandText = $"SELECT max({columName}) FROM {tableName}";

                SqlDataReader reader = com.ExecuteReader();
                if (reader.Read())
                    value = reader.GetInt32(0);

                sqlConnection.Close();

                return value;
            }
        }

        public static void Execute(string sqlCmd, string connectionstr)
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectionstr))
            {
                string value = string.Empty;

                SqlCommand com = new SqlCommand()
                {
                    Connection = sqlConnection
                };
                sqlConnection.Open();
                com.CommandText = sqlCmd;
                com.ExecuteNonQuery();
                sqlConnection.Close();
            }
        }

        public static DataSet ComboDataSet(string sqlCmd, string tablename, string connectionstr)
        {
            DataSet ds = new DataSet();
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(connectionstr))
                {
                    SqlDataAdapter da = new SqlDataAdapter(sqlCmd, sqlConnection);

                    sqlConnection.Open();
                    da.Fill(ds, tablename);

                    sqlConnection.Close();
                }
            }
            catch (FileNotFoundException e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

            
            return ds;
        }

        public static int SqlCountTable(string tableName, string ConnectionStr)
        {
            using (SqlConnection sqlConnection = new SqlConnection(ConnectionStr))
            {

                SqlCommand com = new SqlCommand
                {
                    Connection = sqlConnection
                };
                sqlConnection.Open();

                com.CommandText = "SELECT COUNT(*) FROM " + tableName;

                Int32 count = (Int32)com.ExecuteScalar();

                sqlConnection.Close();

                return count;
            }
        }


    }
}
