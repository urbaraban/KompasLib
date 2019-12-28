using Kompas6API7;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Controls;

namespace KompasLib.Tools
{
    public class KmpsVariable
    {
        private IKompasDocument2D1 doc71;

        public KmpsVariable(IKompasDocument2D1 Doc71)
        {
            this.doc71 = Doc71;
        }

        //Сумма переменных
        public double Sum(string name, ItemCollection itemcolletion)
        {
            double Summ = 0;
            for (int i = 0; i < itemcolletion.Count; i++)
                Summ += Give(name, itemcolletion[i].ToString());
            return Summ;
        }

        //удаляем переменные
        public void Clear(string num)
        {
            DataSet ds = comboDataSet("SELECT Name, InText FROM dbo.Variable WHERE Base='False'", "Variable");
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                doc71.Variable[false, ds.Tables[0].Rows[i]["Name"] + num].Delete();
        }

        private DataSet comboDataSet(string sqlCmd, string TableName)
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
        public void Update(string name, double value, string index, bool dim = false)
        {
            if (doc71 != null)
                if (!doc71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = doc71.Variable[false, name + index];
                    if (var != null) var.Value = value;
                    doc71.UpdateVariables();
                }
        }
        //Добавляем в переменную
        public void Add(string name, double value, string index)
        {
            if (doc71 != null)
                if (!doc71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = doc71.Variable[false, name + index];
                    var.Value += value;
                    doc71.UpdateVariables();
                }
        }

        //Записываем комментарий к переменной
        public void UpdateNote(string name, string value, string index)
        {
            if (doc71 != null)
                if (!doc71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = doc71.Variable[false, name + index];
                    var.Note = value;
                    doc71.UpdateVariables();
                }

        }
        //Получить перменную
        public double Give(string name, string index)
        {
            if (doc71 != null)
            {
                if (!doc71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = doc71.Variable[false, name + index];
                     if (var != null) return var.Value;
                }
                return 0;
            }
            else return 0;
        }
        //Получить коммент к переменной
        public string GiveNote(string name, string index)
        {
            if (doc71 != null)
            {
                if (!doc71.IsVariableNameValid(name + index))
                {
                    IVariable7 var = doc71.Variable[false, name + index];
                    return var.Note;
                }
                return string.Empty;
            }
            else return null;
        }
    }
}
