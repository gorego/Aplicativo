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
using System.IO;

namespace Aplicativo
{
    public partial class frmLiquidacion : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string id = "";

        public void cargarLiquidacion()
        {
            string query = "SELECT * FROM Liquidacion WHERE Trabajador = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView1.DataSource = banco;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Año de Liquidación";
        }

        public void cargarEmpleado()
        {
            string query = "SELECT Nombres + ' ' + Apellidos FROM Trabajadores WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                while (myReader.Read())
                {
                    this.Text = myReader.GetString(0);
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

        public void agregarLiquidacion()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Liquidacion(ano,Trabajador) VALUES (@ano,@Trabajador)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string liquidacion = txtArchivo.Text;
                if (!liquidacion.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + this.Text + "\\Liquidacion\\" + textBox1.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + this.Text + "\\Liquidacion\\" + textBox1.Text, "Liquidacion*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(liquidacion, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + this.Text + "\\Liquidacion\\" + textBox1.Text);

                        string ext = Path.GetExtension(liquidacion);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + this.Text + "\\Liquidacion\\" + textBox1.Text + "\\Liquidacion" + textBox1.Text + ext));
                    }

                    cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = textBox1.Text;
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = id;

                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Liquidación agregada.");
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
                    MessageBox.Show("Favor agregar el recibo de liquidación.", "Error");
                }
            }
            else
            {
                MessageBox.Show("Connection Failed");
            }
        }

        public frmLiquidacion(string trabajador)
        {
            InitializeComponent();
            id = trabajador;
            cargarLiquidacion();
            cargarEmpleado();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void btnPredial_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            txtArchivo.Text = openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarLiquidacion();
            cargarLiquidacion();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            linkLabel1.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + this.Text + "\\Liquidacion\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + this.Text + "\\Liquidacion\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "Liquidacion*");
            if (prueba.Length > 0)
            {
                if (File.Exists(prueba[0]))
                {
                    System.Diagnostics.Process.Start(prueba[0]);
                }
            }
            else
            {
                MessageBox.Show("No se encuentra el archivo.", "Error");
            }
        }
    }
}
