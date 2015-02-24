using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Aplicativo
{
    public partial class frmPluviometria : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmPluviometria()
        {
            InitializeComponent();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd/MM/yyyy";
            Variables.cargar(dataGridView1, "SELECT p.ID, p.numBase, p.codBase, b.Predio FROM BancoTierras AS b INNER JOIN Pluviometria AS p ON b.ID = p.Ubicacion;");
            Variables.cargar(comboBox1, "SELECT * FROM BancoTierras", "Predio");
            dataGridView1.Columns[1].HeaderText = "# de Base";
            dataGridView1.Columns[2].HeaderText = "Codigo";
            dataGridView1.Columns[3].HeaderText = "Ubicación";
            comboBox1.SelectedItem = null;
        }

        public void cargar(ComboBox combo, string query, string display)
        {
            //Ejecutar el query y llenar el ComboBox.
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

        public void cargar(DataGridView data, string query)
        {            
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            data.DataSource = supervisores;
            data.Columns[0].Visible = false;
        }

        public void eliminar(DataGridView data, string query, string notificacion, string mensaje) {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(mensaje, "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = data.Rows[data.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand(query + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show(notificacion);
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
        }

        public void ejecutarQuery(string query,string notificacion)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@numBase", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@codBase", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show(notificacion);
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

        public void ejecutarQueryRegistro(DataGridView data,string query, string notificacion)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Base", OleDbType.VarChar).Value = data.Rows[data.CurrentCell.RowIndex].Cells[0].Value.ToString();
                cmd.Parameters.Add("@Grado", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show(notificacion);
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

        public void reiniciarTablero() 
        {
            textBox1.Text = "";
            textBox2.Text = "";
            comboBox1.Text = "";
            textBox4.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el # de la base", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox2.Text.Equals(""))
            {
                MessageBox.Show("Favor el codigo de la base.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (comboBox1.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar la ubicacion de la base.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                ejecutarQuery("INSERT INTO Pluviometria (numBase,codBase,Ubicacion) VALUES (@numBase,@codBase,@Ubicacion)", "Base agregada.");
                cargar(dataGridView1, "SELECT * FROM Pluviometria");
            }            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ejecutarQuery("UPDATE Pluviometria SET numBase=@numBase,codBase=@codBase,Ubicacion=@Ubicacion WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "Base modificada.");
            cargar(dataGridView1, "SELECT * FROM Pluviometria");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            eliminar(dataGridView1, "DELETE FROM Pluviometria WHERE ID = ", "Base eliminada", "Seguro que desea eliminar la base?");
            cargar(dataGridView1, "SELECT * FROM Pluviometria");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ejecutarQueryRegistro(dataGridView1,"INSERT INTO registroPluviometria (Pluviometria,Grado,Fecha) VALUES (@Pluviometria,@Grado,@Fecha)", "Lluvia registrada.");
            cargar(dataGridView2, "SELECT r.ID, p.codBase, r.Grado, r.Fecha FROM Pluviometria AS p INNER JOIN registroPluviometria AS r ON p.ID = r.Pluviometria WHERE r.Pluviometria = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " ORDER BY r.ID Desc");
            dataGridView2.Columns[1].HeaderText = "Codigo de Base";
            dataGridView2.Columns[2].HeaderText = "mm de Lluvia";
            label8.Text = "Total " + getTotal(dataGridView2) + " mm de Lluvia";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cargar(dataGridView2, "SELECT r.ID, p.codBase, r.Grado, r.Fecha FROM Pluviometria AS p INNER JOIN registroPluviometria AS r ON p.ID = r.Pluviometria WHERE r.Pluviometria = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " ORDER BY r.ID Desc");
            dataGridView2.Columns[1].HeaderText = "Codigo de Base";
            dataGridView2.Columns[2].HeaderText = "mm de Lluvia";
            dataGridView2.Columns[1].Visible = false;
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            comboBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            label8.Text = "Total " + getTotal(dataGridView2) + " mm de Lluvia";
        }

        public double getTotal(DataGridView data)
        {
            double total = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                total += double.Parse(data.Rows[i].Cells[2].Value.ToString());
            }
            return total;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            eliminar(dataGridView2, "DELETE FROM registroPluviometria WHERE ID = ", "Lluvia eliminada", "Seguro que desea eliminar la lluvia registrada?");
            cargar(dataGridView2, "SELECT r.ID, p.codBase, r.Grado, r.Fecha FROM Pluviometria AS p INNER JOIN registroPluviometria AS r ON p.ID = r.Pluviometria WHERE r.Pluviometria = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " ORDER BY r.ID Desc");
            dataGridView2.Columns[1].HeaderText = "Codigo de Base";
            dataGridView2.Columns[2].HeaderText = "mm de Lluvia";
            label8.Text = "Total " + getTotal(dataGridView2) + " mm de Lluvia";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            IFormatProvider culture = new System.Globalization.CultureInfo("es-CO", true);
            DateTime inicio = dateTimePicker2.Value;
            DateTime final = dateTimePicker3.Value;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                DateTime date = DateTime.Parse(dataGridView2.Rows[i].Cells[3].Value.ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                if (!enRango(inicio, final, date))
                {
                    dataGridView2.Rows.RemoveAt(i);
                    i--;
                }                   
            }
            label8.Text = "Total " + getTotal(dataGridView2) + " mm de Lluvia";
        }

        public bool enRango(DateTime inicio, DateTime final, DateTime date)
        {
            if (date <= final && date >= inicio)
                return true;
            else
                return false;
        }

        public void imprimirPluviometria(DataGridView data)
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            imprimirPluviometria(dataGridView1);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            imprimirPluviometria(dataGridView2);
        }
    }
}
