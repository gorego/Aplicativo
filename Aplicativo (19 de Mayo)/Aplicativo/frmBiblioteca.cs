﻿using System;
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
    public partial class frmBiblioteca : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int carpetaID = 0;

        public frmBiblioteca()
        {
            InitializeComponent();
            Variables.cargar(dataGridView1, "SELECT * FROM Biblioteca");
            Variables.cargar(dataGridView2, "SELECT * FROM Carpetas");            
            //cargarBiblioteca();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        public void agregarBiblioteca()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Biblioteca (Nombre,Descripcion,Ubicacion,Tipo) VALUES (@Nombre,@Descripcion,@Ubicacion,@Tipo)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string archivo = textBox4.Text;
                if (!archivo.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca");
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca", textBox1.Text + "*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(archivo, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca");
                        string ext = Path.GetExtension(archivo);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca\\" + textBox1.Text + ext));
                        fs.Close();
                    }
                }
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = textBox5.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Archivo agregado.");
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

        public void agregarCarpeta()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Carpetas (Carpeta,Ubicacion,Descripcion) VALUES (@Carpeta,@Ubicacion,@Descripcion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca\\" + textBox9.Text);
                cmd.Parameters.Add("@Carpeta", OleDbType.VarChar).Value = textBox9.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox7.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = textBox8.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Carpeta agregada.");
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

        public void cargarBiblioteca()
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM Biblioteca";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int i = 0;
            int j = 1;
            try
            {
                while (myReader.Read())
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = j++;
                    dataGridView1.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView1.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView1.Rows[i].Cells[3].Value = myReader.GetString(3);
                    dataGridView1.Rows[i].Cells[4].Value = myReader.GetString(4);     
                    i++;
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar nombre del archivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox5.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el tipo de archivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarBiblioteca();
                Variables.cargar(dataGridView1, "SELECT * FROM Biblioteca");
                //cargarBiblioteca();
                textBox1.Text = "";
                textBox2.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Biblioteca WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Archivo eliminado.");
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca");
                        string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca", textBox1.Text + "*");
                        if (prueba.Length > 0)
                        {
                            if (File.Exists(prueba[0]))
                            {

                                File.Delete(prueba[0]);
                            }
                        }
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
                cargarBiblioteca();
            }
        }

        public void modificarBiblioteca()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Biblioteca SET Nombre=@Nombre,Descripcion=@Descripcion,Ubicacion=@Ubicacion,Tipo=@Tipo WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string contrato = textBox4.Text;
                if (!contrato.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca");
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca", textBox1.Text + "*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(contrato, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca");
                        string ext = Path.GetExtension(contrato);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca\\" + textBox1.Text + ext));
                    }
                }
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = textBox5.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Formato modificado.");
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Biblioteca\\", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "*");
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
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
        }

        public void buscarBiblioteca()
        {
            //while (dataGridView1.Rows.Count != 0)
            //{
            //    dataGridView1.Rows.RemoveAt(0);
            //}
            string query = "SELECT * FROM Biblioteca ";
            int i = 0;
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Nombre LIKE '%" + textBox1.Text + "%'";
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Descripcion LIKE '%" + textBox2.Text + "%'";
            }
            if (!textBox3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Ubicacion LIKE '%" + textBox3.Text + "%'";
            }
            if (!textBox5.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Tipo LIKE '%" + textBox5.Text + "%'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
            ////Ejecutar el query y llenar el GridView.
            //conn.ConnectionString = connectionString;
            //OleDbCommand cmd = new OleDbCommand(query, conn);
            //cmd.Connection = conn;
            //conn.Open();
            //OleDbDataReader myReader = cmd.ExecuteReader();
            //i = 0;
            //try
            //{
            //    while (myReader.Read())
            //    {
            //        dataGridView1.Rows.Add();
            //        dataGridView1.Rows[i].Cells[0].Value = myReader.GetInt32(0);
            //        dataGridView1.Rows[i].Cells[1].Value = myReader.GetString(1);
            //        dataGridView1.Rows[i].Cells[2].Value = myReader.GetString(2);
            //        dataGridView1.Rows[i].Cells[3].Value = myReader.GetString(3);
            //        i++;
            //    }
            //}
            //finally
            //{
            //    // always call Close when done reading.
            //    myReader.Close();
            //    // always call Close when done reading.
            //    conn.Close();
            //}
        }

        private void button4_Click(object sender, EventArgs e)
        {
            buscarBiblioteca();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            modificarBiblioteca();
            cargarBiblioteca();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar nombre de la carpeta.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarCarpeta();
                Variables.cargar(dataGridView2, "SELECT * FROM Carpetas");
                //cargarBiblioteca();
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            tabPage2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString();
            carpetaID = Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());            
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            tabPage2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString();
            carpetaID = Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());
        }
    }
}
