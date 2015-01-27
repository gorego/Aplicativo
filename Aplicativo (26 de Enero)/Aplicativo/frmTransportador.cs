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
    public partial class frmTransportador : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string id = "";
        public frmTransportador(string proveedor)
        {
            InitializeComponent();
            id = proveedor;            
            cargarTransportadores();
        }

        public void cargarTransportadores()
        {
            string query = "SELECT ID, (Nombres + ' ' + Apellidos) As Nombre , Cedula, Licencia, Placa FROM Transportadores WHERE Nombres <> 'N/A' AND Proveedor = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;            
        }

        public void subirArchivo(string archivo, TextBox text) 
        {
            string arch = text.Text;
            if (!arch.Equals("")) 
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text + " " + textBox2.Text);
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text + " " + textBox2.Text, archivo+"*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(arch, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text + " " + textBox2.Text);
                    string ext = Path.GetExtension(arch);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text + " " + textBox2.Text + "\\"+ archivo + ext));
                }
            }
        }

        public void agregarTransportador()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Transportadores (Nombres,Apellidos,Cedula,Licencia,Placa,Proveedor) VALUES (@Nombres,@Apellidos,@Cedula,@Licencia,@Placa,@Proveedor)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                subirArchivo("Cedula", textBox6);
                subirArchivo("Licencia", textBox7);
                subirArchivo("Placa", textBox8);                
                cmd.Parameters.Add("@Nombres", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Apellidos", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Licencia", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = textBox5.Text;
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = id;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Transportador agregado.");
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

        public void modificarTransportador()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Transportadores SET Nombres=@Nombres,Apellidos=@Apellidos,Cedula=@Cedula,Licencia=@Licencia,Placa=@Placa WHERE Proveedor = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                subirArchivo("Cedula", textBox6);
                subirArchivo("Licencia", textBox7);
                subirArchivo("Placa", textBox8);                
                cmd.Parameters.Add("@Nombres", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Apellidos", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Licencia", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = textBox5.Text;
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = id;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Transportador modificado.");
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

        public void eliminarTranpostador() 
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el transportador " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Transportadores WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Transportador eliminado.");
                        //Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text);
                        //if (Directory.Exists("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text + " " + textBox2.Text))
                        //{
                        //    Directory.Delete("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + textBox1.Text + " " + textBox2.Text);
                        //}
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

        public void buscarTransportador()
        {
            string query = "SELECT ID, (Nombres + ' ' + Apellidos) As Nombre , Cedula, Licencia, Placa FROM Transportadores ";
            int i = 0;
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Nombres LIKE '%" + textBox1.Text + "%'";
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Apellidos LIKE '%" + textBox2.Text + "%'";
            }
            if (!textBox3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Cedula LIKE '%" + textBox3.Text + "%'";
            }
            if (!textBox4.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Licencia LIKE '%" + textBox4.Text + "%'";
            }
            if (!textBox5.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Placa LIKE '%" + textBox5.Text + "%'";
            }
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            abrirArchivo("Licencia");
        }

        public void abrirArchivo(string archivo) 
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Transportadores\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), archivo+"*");
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            abrirArchivo("Cedula");
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            abrirArchivo("Placa");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarTransportador();
            cargarTransportadores();
            reiniciarTablero();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox6.Text = openFileDialog1.FileName;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox7.Text = openFileDialog1.FileName;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox8.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            modificarTransportador();
            cargarTransportadores();
            reiniciarTablero();
        }

        public void reiniciarTablero() 
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";   
        }

        private void button5_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            eliminarTranpostador();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            buscarTransportador();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string[] nombre = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString().Split(' ');
            textBox1.Text = nombre[0];
            textBox2.Text = nombre[1];
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
