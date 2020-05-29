using Kompas6API7;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace KompasLib.Tools
{
    public static class KVariable
    {
        //Сумма переменных
        public static double Sum(string name, ItemCollection itemcolletion)
        {
            double Summ = 0;
            for (int i = 0; i < itemcolletion.Count; i++)
                Summ += Give(name, itemcolletion[i].ToString());
            return Summ;
        }

        //удаляем переменные
        public static async Task ClearAsync(string num, DataTable variableTable)
        {
           await Task.Factory.StartNew(() =>
            {
                if (variableTable == null)
                    variableTable = comboDataSet("SELECT Name, InText FROM dbo.Variable WHERE Base='False'", "Variable").Tables[0];

                DataRow[] rows = variableTable.Select("Base = false");

                for (int i = 0; i < rows.Length; i++)
                    if (!KmpsDoc.D71.IsVariableNameValid(rows[i]["Name"] + num))
                        KmpsDoc.D71.Variable[false, rows[i]["Name"] + num].Delete();
            });

        }

        private static DataSet comboDataSet(string sqlCmd, string TableName)
        {
            DataSet ds = new DataSet();

            string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\DefaultValue.mdf;Integrated Security=True";

            using (SqlConnection sqlConnection = new SqlConnection(sqlCon))
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlCmd, sqlConnection);
                sqlConnection.Open();
                da.Fill(ds, TableName);

                sqlConnection.Close();
            }
            return ds;
        }

        //Записываем значение переменной
        public static async Task<bool> UpdateAsync(string name, double value, string index, bool dim = false)
        {
            if (KmpsDoc.D71 != null)
                if (!KmpsDoc.D71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = KmpsDoc.D71.Variable[false, name + index];
                    if (var != null) var.Value = value;
                    KmpsDoc.D71.UpdateVariables();
                }
            return true;
        }
        //Добавляем в переменную
        public static async Task<bool> Add(string name, double value, string index)
        {
                if (KmpsDoc.D71 != null)
                    if (!KmpsDoc.D71.IsVariableNameValid(name + index))
                    {
                        IVariable7 var = KmpsDoc.D71.Variable[false, name + index];
                        var.Value += value;
                        KmpsDoc.D71.UpdateVariables();
                    }
                return true;
        }

        //Записываем комментарий к переменной
        public static async Task UpdateNote(string name, string value, string index)
        {
            await Task.Factory.StartNew(() =>
            {
                if (KmpsDoc.D71 != null)
                    if (!KmpsDoc.D71.IsVariableNameValid(name + index))
                    {
                        IVariable7 var = KmpsDoc.D71.Variable[false, name + index];
                        if (var != null)
                            var.Note = value;
                        KmpsDoc.D71.UpdateVariables();
                    }
            });

        }
        //Получить перменную
        public static double Give(string name, string index)
        {
            if (KmpsDoc.D71 != null)
            {
                if (!KmpsDoc.D71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = KmpsDoc.D71.Variable[false, name + index];
                     if (var != null) return var.Value;
                }
                return 0;
            }
            else return 0;
        }
        //Получить коммент к переменной
        public static string GiveNote(string name, string index)
        {
            if (KmpsDoc.D71 != null)
            {
                if (!KmpsDoc.D71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = KmpsDoc.D71.Variable[false, name + index];
                    return var.Note;
                }
                return string.Empty;
            }
            else return null;
        }
    }
}
