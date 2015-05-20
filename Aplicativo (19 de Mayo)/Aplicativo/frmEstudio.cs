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
    public partial class frmEstudio : Form
    {
        string lote = "";
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarEstudio(int ano)
        {
            string query = "SELECT * FROM Estudio WHERE lote = " + lote + " and ano = '" + ano + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinarias = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(maquinarias);
            dataGridView1.DataSource = maquinarias;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[17].Visible = false;
        }

        public void reiniciarTablero()
        {
            while (dataGridView2.Rows.Count != 0)
            {
                dataGridView2.Rows.RemoveAt(0);
            }
            dataGridView2.Rows.Add();
            dataGridView2.Rows[0].Cells[0].Value = "30cm";
            dataGridView2.Rows.Add();
            dataGridView2.Rows[1].Cells[0].Value = "60cm";
            dataGridView2.Rows.Add();
            dataGridView2.Rows[2].Cells[0].Value = "90cm";
        }

        public void cargarEstudioActual(int ano)
        {
            string query = "SELECT * FROM Estudio WHERE lote = " + lote + " and ano = '" + ano + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int i = 0;
            try
            {
                while (myReader.Read())
                {
                    dataGridView2.Rows.Add();
                    dataGridView2.Rows[i].Cells[0].Value = myReader.GetString(2);
                    dataGridView2.Rows[i].Cells[1].Value = myReader.GetString(3);
                    dataGridView2.Rows[i].Cells[2].Value = myReader.GetString(4);
                    dataGridView2.Rows[i].Cells[3].Value = myReader.GetString(5);
                    dataGridView2.Rows[i].Cells[4].Value = myReader.GetString(6);
                    dataGridView2.Rows[i].Cells[5].Value = myReader.GetString(7);
                    dataGridView2.Rows[i].Cells[6].Value = myReader.GetString(8);
                    dataGridView2.Rows[i].Cells[7].Value = myReader.GetString(9);
                    dataGridView2.Rows[i].Cells[8].Value = myReader.GetString(10);
                    dataGridView2.Rows[i].Cells[9].Value = myReader.GetString(11);
                    dataGridView2.Rows[i].Cells[10].Value = myReader.GetString(12);
                    dataGridView2.Rows[i].Cells[11].Value = myReader.GetString(13);
                    dataGridView2.Rows[i].Cells[12].Value = myReader.GetString(14);
                    dataGridView2.Rows[i].Cells[13].Value = myReader.GetString(15);
                    dataGridView2.Rows[i].Cells[14].Value = myReader.GetString(16);
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

        public void cargarIMA()
        {
            while (dataGridView4.Rows.Count != 0)
            {
                dataGridView4.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM IMA WHERE Lote = " + lote + " ORDER BY ID Desc";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int i = 0;
            try
            {
                while (myReader.Read())
                {
                    dataGridView4.Rows.Add();
                    dataGridView4.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView4.Rows[i].Cells[1].Value = myReader.GetInt32(1);
                    dataGridView4.Rows[i].Cells[2].Value = myReader.GetDouble(2);
                    dataGridView4.Rows[i].Cells[3].Value = myReader.GetDouble(3);
                    dataGridView4.Rows[i].Cells[4].Value = myReader.GetInt32(4);
                    dataGridView4.Rows[i].Cells[5].Value = myReader.GetDouble(5);
                    dataGridView4.Rows[i].Cells[6].Value = myReader.GetDouble(6);
                    dataGridView4.Rows[i].Cells[7].Value = myReader.GetString(7);
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

        public int getID() {
            string query = "SELECT id FROM Lotes WHERE Codigo = "+lote+"";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            try
            {
                while (myReader.Read())
                {
                    id = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return id;
        }

        public void cargarViasActual(int id)
        {
            string query = "SELECT * FROM ViasDeMovilizacion WHERE Lote = '" + id + "'";
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
                    txt1.Text = myReader.GetInt32(1).ToString();
                    txt2.Text = myReader.GetInt32(2).ToString();
                    txt3.Text = myReader.GetInt32(3).ToString();
                    txt4.Text = myReader.GetInt32(4).ToString();
                    txt5.Text = myReader.GetInt32(5).ToString();
                    txt6.Text = myReader.GetInt32(6).ToString();
                    txt7.Text = myReader.GetInt32(7).ToString();          
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

        public void cargarAno() {
            string query = "SELECT ano FROM Estudio WHERE lote = " + lote + " GROUP BY ano";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinarias = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(maquinarias);
            dataGridView3.DataSource = maquinarias;
            dataGridView3.Columns[0].HeaderText = "Año de Estudio";
        }

        public void agregarEstudio(int perfil) {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Estudio (ano,Perfil,PH,CO,N,Ca,K,Mg,Na,Al,CICE,P,Ar,L,A,Textura,Lote) VALUES (@ano,@Perfil,@PH,@CO,@N,@Ca,@K,@Mg,@Na,@Al,@CICE,@P,@Ar,@L,@A,@Textura,@Lote)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Perfil", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[0].Value;
                cmd.Parameters.Add("@PH", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[1].Value;
                cmd.Parameters.Add("@CO", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[2].Value;
                cmd.Parameters.Add("@N", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[3].Value;
                cmd.Parameters.Add("@Ca", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[4].Value;
                cmd.Parameters.Add("@K", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[5].Value;
                cmd.Parameters.Add("@Mg", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[6].Value;
                cmd.Parameters.Add("@Na", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[7].Value;
                cmd.Parameters.Add("@Al", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[8].Value;
                cmd.Parameters.Add("@CICE", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[9].Value;
                cmd.Parameters.Add("@P", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[10].Value;
                cmd.Parameters.Add("@Ar", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[11].Value;
                cmd.Parameters.Add("@L", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[12].Value;
                cmd.Parameters.Add("@A", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[13].Value;
                cmd.Parameters.Add("@Textura", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[14].Value;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = lote;

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

        public void modificarEstudio(int perfil)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("Update Estudio SET Perfil=@Perfil,PH=@PH,CO=@CO,N=@N,Ca=@Ca,K=@K,Mg=@Mg,Na=@Na,Al=@Al,CICE=@CICE,P=@P,Ar=@Ar,L=@L,A=@A,Textura=@Textura WHERE ano = '" + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString() + "' AND ID = " + dataGridView1.Rows[perfil].Cells[0].Value);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Perfil", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[0].Value;
                cmd.Parameters.Add("@PH", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[1].Value;
                cmd.Parameters.Add("@CO", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[2].Value;
                cmd.Parameters.Add("@N", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[3].Value;
                cmd.Parameters.Add("@Ca", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[4].Value;
                cmd.Parameters.Add("@K", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[5].Value;
                cmd.Parameters.Add("@Mg", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[6].Value;
                cmd.Parameters.Add("@Na", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[7].Value;
                cmd.Parameters.Add("@Al", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[8].Value;
                cmd.Parameters.Add("@CICE", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[9].Value;
                cmd.Parameters.Add("@P", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[10].Value;
                cmd.Parameters.Add("@Ar", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[11].Value;
                cmd.Parameters.Add("@L", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[12].Value;
                cmd.Parameters.Add("@A", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[13].Value;
                cmd.Parameters.Add("@Textura", OleDbType.VarChar).Value = dataGridView2.Rows[perfil].Cells[14].Value;
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

        public void modificarVia(int cod)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("Update ViasDeMovilizacion SET distanciaAserradero=@distanciaAserradero,arroyo=@arroyo,criticos=@criticos,portones=@portones,tiempo=@tiempo,distanciaBarranquilla=@distanciaBarranquilla,puentes=@puentes WHERE Lote = '" + cod + "'");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@distanciaAserradero", OleDbType.VarChar).Value = txt1.Text;
                cmd.Parameters.Add("@arroyo", OleDbType.VarChar).Value = txt2.Text;
                cmd.Parameters.Add("@criticos", OleDbType.VarChar).Value = txt3.Text;
                cmd.Parameters.Add("@portones", OleDbType.VarChar).Value = txt4.Text;
                cmd.Parameters.Add("@tiempo", OleDbType.VarChar).Value = txt5.Text;
                cmd.Parameters.Add("@distanciaBarranquilla", OleDbType.VarChar).Value = txt6.Text;
                cmd.Parameters.Add("@puentes", OleDbType.VarChar).Value = txt7.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Via modificada.");
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

        public void cargarPredio()
        {
            string query = "SELECT * FROM Lotes WHERE Codigo = " + lote;
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


        public frmEstudio(string id)
        {
            InitializeComponent();
            lote = id;
            cargarAno();
            cargarPredio();
            dataGridView2.Rows.Add();
            dataGridView2.Rows[0].Cells[0].Value = "30cm";
            dataGridView2.Rows.Add();
            dataGridView2.Rows[1].Cells[0].Value = "60cm";
            dataGridView2.Rows.Add();
            dataGridView2.Rows[2].Cells[0].Value = "90cm";
            int cod = getID();
            cargarViasActual(cod);
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            cargarIMA();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarEstudio(0);
            agregarEstudio(1);
            agregarEstudio(2);
            cargarAno();
            reiniciarTablero();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            modificarEstudio(0);
            modificarEstudio(1);
            modificarEstudio(2);
            cargarEstudio(Int32.Parse(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            cargarAno();
            reiniciarTablero();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar estudio del año " + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string ano = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Estudio WHERE ano = '" + ano + "' and Lote = " + lote);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Estudio eliminado.");
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
                cargarEstudio(Int32.Parse(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString()));
                cargarAno();
                reiniciarTablero();
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cargarEstudio(Int32.Parse(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            textBox1.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
            cargarEstudioActual(Int32.Parse(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString()));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int cod = getID();
            modificarVia(cod);
            cargarViasActual(cod);

            string PDF = textBox7.Text;
            if (!PDF.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis", "CroquisPDF*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(PDF, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(PDF);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis" + "\\CroquisPDF" + ext));
                }
            }
            string SHP = textBox6.Text;
            if (!SHP.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis", "CroquisSHP*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(SHP, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis" + "\\CroquisSHP" + ext));
                }
            }
            string KMZ = textBox3.Text;
            if (!KMZ.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\KMZ");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\KMZ", "KMZ*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(KMZ, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\KMZ");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\KMZ" + "\\KMZ" + ext));
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string PDF = textBox5.Text;
            if (!PDF.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis", "CroquisPDF*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(PDF, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(PDF);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis" + "\\CroquisPDF" + ext));
                }
            }
            string SHP = textBox2.Text;
            if (!SHP.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis", "CroquisSHP*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(SHP, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis" + "\\CroquisSHP" + ext));
                }
            }
            string KMZ = textBox4.Text;
            if (!KMZ.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\KMZ");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\KMZ", "KMZ*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(KMZ, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\KMZ");

                    string ext = Path.GetExtension(KMZ);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\KMZ" + "\\KMZ" + ext));
                }
            }
            string Registro = textBox8.Text;
            if (!Registro.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Registro");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Registro", "Registro*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(Registro, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Registro");

                    string ext = Path.GetExtension(Registro);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Registro" + "\\Registro" + ext));
                }
            }
            MessageBox.Show("Archivos modificados.");
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis", "CroquisPDF*");
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
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Croquis", "CroquisSHP*");
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
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\KMZ");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\KMZ", "KMZ*");
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

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox7.Text = openFileDialog1.FileName;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox6.Text = openFileDialog1.FileName;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox3.Text = openFileDialog1.FileName;
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis", "CroquisPDF*");
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

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\Croquis", "CroquisSHP*");
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

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\KMZ");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Vias\\" + this.Text + "\\KMZ", "KMZ*");
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

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox8.Text = openFileDialog1.FileName;
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Registro");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Lotes\\" + this.Text + "\\Registro", "Registro*");
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

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void textBox12_Enter(object sender, EventArgs e)
        {
            textBox12.Text = "";
        }

        private void textBox13_Enter(object sender, EventArgs e)
        {
            IFormatProvider culture = new System.Globalization.CultureInfo("es-CO", true);
            if (dataGridView4.Rows.Count > 0)
            {
                DateTime d1 = DateTime.Parse(dataGridView4.Rows[0].Cells[7].Value.ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                DateTime d2 = dateTimePicker2.Value;
                double IMA = (double.Parse(textBox12.Text) - double.Parse(dataGridView4.Rows[0].Cells[5].Value.ToString()))/(((d2-d1).TotalDays/365));
                if (IMA.ToString().Length >= 6)
                    textBox13.Text = IMA.ToString().Substring(0, 6);
                else
                    textBox13.Text = IMA.ToString();
            }
            else
            {
                double IMA = (double.Parse(textBox12.Text));
                if (IMA.ToString().Length >= 6)
                    textBox13.Text = IMA.ToString().Substring(0, 6);
                else
                    textBox13.Text = IMA.ToString();
            }
        }

        public void agregarIMA()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO IMA (Lote,Diametro,Altura,Cantidad,Volumen,IMA,Fecha) VALUES (@Lote,@Diametro,@Altura,@Cantidad,@Volumen,@IMA,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = lote;
                cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox9.Text;
                cmd.Parameters.Add("@Altura", OleDbType.VarChar).Value = textBox10.Text;
                cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox11.Text;
                cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = textBox12.Text;
                cmd.Parameters.Add("@IMA", OleDbType.VarChar).Value = textBox13.Text;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("dd") + "/" + dateTimePicker2.Value.ToString("MM") + "/" + dateTimePicker2.Value.Year;
                try
                {
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("IMA registrado.");
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

        private void button12_Click(object sender, EventArgs e)
        {
            agregarIMA();
            cargarIMA();
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            textBox9.Text = "";
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            textBox10.Text = "";
        }

        private void textBox11_Enter(object sender, EventArgs e)
        {
            textBox11.Text = "";
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (textBox9.Text.Equals(""))
                textBox9.Text = "0";
        }

        private void textBox10_Leave(object sender, EventArgs e)
        {
            if (textBox10.Text.Equals(""))
                textBox10.Text = "0";
        }

        private void textBox11_Leave(object sender, EventArgs e)
        {
            if (textBox11.Text.Equals(""))
                textBox11.Text = "0";
        }

        private void textBox12_Leave(object sender, EventArgs e)
        {
            if (textBox12.Text.Equals(""))
                textBox12.Text = "0";
        }

        private void button13_Click(object sender, EventArgs e)
        {
             DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar inventario de la fecha " + dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[7].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

             if (dialogResult == DialogResult.Yes)
             {

                 string ano = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
                 conn.ConnectionString = connectionString;
                 OleDbCommand cmd = new OleDbCommand("DELETE FROM IMA WHERE id = " + dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString());
                 cmd.Connection = conn;
                 conn.Open();

                 if (conn.State == ConnectionState.Open)
                 {
                     try
                     {
                         cmd.ExecuteNonQuery();
                         MessageBox.Show("Inventario eliminado.");
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
                 cargarIMA();
             }
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView4);
        }

    }
}
