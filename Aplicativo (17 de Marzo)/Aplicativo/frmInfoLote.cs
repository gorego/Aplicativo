using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace Aplicativo
{
    public partial class frmInfoLote : Form
    {
        string lote = "";
        string opc = "";
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        public frmInfoLote(string id, string tipo)
        {
            InitializeComponent();
            lote = id;
            opc = tipo;
            cargarPredio();
            if (!tipo.Equals("LoteGanadero"))
                this.Width = 383;
            else
            {
                Variables.cargar(dataGridView3, "SELECT * FROM loteHeno WHERE Lote = " + id + " ORDER BY ID Desc");
                dataGridView3.Columns[1].Visible = false;
                dataGridView3.Columns[2].HeaderText = "Año";
                dataGridView3.Columns[3].HeaderText = "# de Bolas";
            }
        }

        public void cargarPredio()
        {
            string query = "SELECT * FROM " + opc + " WHERE Codigo = " + lote;
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
                    this.Text = myReader.GetString(2);
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

        private void btnPDF_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox5.Text = openFileDialog1.FileName;
        }

        private void btnSHP_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox2.Text = openFileDialog1.FileName;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string PDF = textBox5.Text;
            if (!PDF.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis", "CroquisPDF*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(PDF, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(PDF);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis" + "\\CroquisPDF" + ext));
                }
            }
            string SHP = textBox2.Text;
            if (!SHP.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis", "CroquisSHP*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(SHP, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis" + "\\CroquisSHP" + ext));
                }
            }
            string KMZ = textBox4.Text;
            if (!KMZ.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\KMZ");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\KMZ", "KMZ*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(KMZ, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\KMZ");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\KMZ" + "\\KMZ" + ext));
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis", "CroquisPDF*");
            if (prueba.Length > 0)
            {
                if (File.Exists(prueba[0]))
                {
                    System.Diagnostics.Process.Start(prueba[0]);
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo.", "Error");
                }
            }
            else
            {
                MessageBox.Show("No se encuentra el archivo.", "Error");
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\Croquis", "CroquisSHP*");
            if (prueba.Length > 0)
            {
                if (File.Exists(prueba[0]))
                {
                    System.Diagnostics.Process.Start(prueba[0]);
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo.", "Error");
                }
            }
            else
            {
                MessageBox.Show("No se encuentra el archivo.", "Error");
            }
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\KMZ");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Lotes\\" + this.Text + "\\KMZ", "KMZ*");
            if (prueba.Length > 0)
            {
                if (File.Exists(prueba[0]))
                {
                    System.Diagnostics.Process.Start(prueba[0]);
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo.", "Error");
                }
            }
            else
            {
                MessageBox.Show("No se encuentra el archivo.", "Error");
            }
        }

        public void agregarHeno()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO loteHeno (Lote,Ano,Bolas) VALUES (@Lote,@Ano,@Bolas)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = lote;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Bolas", OleDbType.VarChar).Value = textBox3.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Agregado.");
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                if (!textBox3.Text.Equals(""))
                {
                    agregarHeno();
                    Variables.cargar(dataGridView3, "SELECT * FROM loteHeno WHERE Lote = " + lote);
                    dataGridView3.Columns[1].Visible = false;
                    dataGridView3.Columns[2].HeaderText = "Año";
                    dataGridView3.Columns[3].HeaderText = "# de Bolas";
                    textBox3.Text = "0";
                    textBox1.Text = "";
                }
                else
                {
                    MessageBox.Show("Favor ingresar el numero de bolas de heno.","Error");
                }
            }
            else
            {
                MessageBox.Show("Favor ingresar el año.","Error");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el año " + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[2].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM loteHeno WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Eliminado.");
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
            Variables.cargar(dataGridView3, "SELECT * FROM loteHeno WHERE Lote = " + lote);
            dataGridView3.Columns[1].Visible = false;
            dataGridView3.Columns[2].HeaderText = "Año";
            dataGridView3.Columns[3].HeaderText = "# de Bolas";
            textBox3.Text = "0";
            textBox1.Text = "";
        }
    }
}
