using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Aplicativo
{
    static class Variables
    {
        public static String connectionString =
        @"Provider=Microsoft.ACE.OLEDB.12.0;Data"
        + @" Source=C:\Users\" + Environment.UserName + @"\Dropbox\Aplicativo\BaseDeDatos2.accdb;Jet OLEDB:Database Password=4545g";

        private static string _userName = "";
        private static int _tipo = 0;

        public static string userName
        {
            get { return _userName; }
            set { _userName = value; }
        }

        public static int tipo
        {
            get { return _tipo; }
            set { _tipo = value; }
        }

        private static int _userID = 0;

        public static int userID
        {
            get { return _userID; }
            set { _userID = value; }
        }

        public static void cargar(ComboBox combo, string query, string display)
        {
            //Ejecutar el query y llenar el ComboBox.
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = display;
            combo.ValueMember = "ID";
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public static void imprimir(DataGridView data)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                XcelApp.Application.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 2; i < data.Columns.Count + 1; i++)
                {
                    XcelApp.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 1; j < data.Columns.Count; j++)
                    {
                        XcelApp.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = XcelApp.Range[XcelApp.Cells[i + 2, 3], XcelApp.Cells[i + 2, data.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                excelCellrange = XcelApp.Range[XcelApp.Cells[2, 3], XcelApp.Cells[data.Rows.Count + 2, data.Columns.Count + 1]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
                XcelApp.Columns.AutoFit();
                XcelApp.Visible = true;
            }
        }

        public static void cargar2(ComboBox combo, string query, string display)
        {
            //Ejecutar el query y llenar el ComboBox.
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                while (myReader.Read())
                {
                    combo.Items.Add(myReader.GetString(0));
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();         
            }
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;

        }

        public static void cargar(DataGridViewComboBoxColumn combo, string query, string display, DataGridView data)
        {
            //Ejecutar el query y llenar el ComboBox.
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = display;
            combo.ValueMember = "ID";
        }


        public static void cargar(DataGridView data, string query)
        {
            //Ejecutar el query y llenar el GridView.
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            data.DataSource = supervisores;
            data.Columns[0].Visible = false;
        }

        public static void agregarLog(string accion, string usuario)
        {
            OleDbConnection conn = new OleDbConnection();           
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoIngresos (Usuario,Accion,Fecha) VALUES (@Usuario,@Accion,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Usuario", OleDbType.VarChar).Value = usuario;
                cmd.Parameters.Add("@Accion", OleDbType.VarChar).Value = accion;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year + " - " + DateTime.Now.Hour + ":" + DateTime.Now.ToString("mm") + ":" + DateTime.Now.ToString("ss");
                try
                {
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show(ex.Source);
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Connection Failed");
            }
        }

        public static bool existe(string query)
        {   
            bool sw = true;
            OleDbConnection conn = new OleDbConnection();           
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        public static void dirtyCell(DataGridView data)
        {
            if (data.IsCurrentCellDirty)
                data.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }
}
