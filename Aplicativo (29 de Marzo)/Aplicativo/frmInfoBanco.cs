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
    public partial class frmInfoBanco : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string id = "";        

        public void cargarPredial() {
            string query = "SELECT * FROM Predial WHERE Predio = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView1.DataSource = banco;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Año Impuesto Predial";
        }

        public void cargarPredio()
        {
            string query = "SELECT * FROM BancoTierras WHERE ID = " + id;
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
                    this.Text = myReader.GetString(1);
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

        public void agregarPredial() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Predial(ano,Predio) VALUES (@ano,@Predio)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string predial = txtPredial.Text;
                if (!predial.Equals(""))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + txtAno.Text);
                    string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + txtAno.Text, "Impuesto*");
                    if (prueba.Length > 0)
                    {
                        if (File.Exists(prueba[0]))
                        {

                            File.Delete(prueba[0]);
                        }
                    }
                    using (FileStream fs = File.Open(predial, FileMode.Open))
                    {
                        Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + txtAno.Text);

                        string ext = Path.GetExtension(predial);
                        fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + txtAno.Text + "\\Impuesto" + txtAno.Text + ext));
                    }

                    cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = txtAno.Text;
                    cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = id;

                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Impuesto predial agregado.");
                        conn.Close();
                    }
                    catch (OleDbException ex)
                    {
                        MessageBox.Show(ex.Source);
                        conn.Close();
                    }
                }
                else {
                    MessageBox.Show("Favor agregar el impuesto predial", "Error");
                }
            }
            else
            {
                MessageBox.Show("Connection Failed");
            }
        }

        public void totalOrdenes()
        {
            int total = 0;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                total += Int32.Parse(dataGridView3.Rows[i].Cells[7].Value.ToString());
            }
            label3.Text = "Total: " + String.Format("{0:c}", total);
        }

        public frmInfoBanco(string banco)
        {
            InitializeComponent();
            id = banco;
            cargarPredio();
            cargarPredial();
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            cargarLotes();
            cargarOrdenes("b.ID", banco);
            totalOrdenes();
        }

        public void cargarOrdenes(string tipo, string id)
        {
            string query = "SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.costoFinal, h.estadoOrden FROM Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN ((BancoTierras AS b INNER JOIN Areas AS area ON b.ID = area.Predio) INNER JOIN historicoOrdenes AS h ON area.Codigo = h.Lote) ON t.ID = h.Supervisor) ON a.ID = h.Actividad WHERE " + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.costoFinal, h.estadoOrden FROM Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN ((BancoTierras AS b INNER JOIN Lotes AS area ON b.ID = area.Predio) INNER JOIN historicoOrdenes AS h ON area.Codigo = h.Lote) ON t.ID = h.Supervisor) ON a.ID = h.Actividad WHERE " + tipo + " = " + id + " UNION ALL SELECT h.ID, h.OT, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.costoFinal, h.estadoOrden FROM Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN ((BancoTierras AS b INNER JOIN LoteGanadero AS area ON b.ID = area.Predio) INNER JOIN historicoOrdenes AS h ON area.Codigo = h.Lote) ON t.ID = h.Supervisor) ON a.ID = h.Actividad  WHERE " + tipo + " = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView3.DataSource = banco;
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[1].HeaderText = "Orden de Trabajo #";
            dataGridView3.Columns[4].HeaderText = "Fecha de Inicio";
            dataGridView3.Columns[5].HeaderText = "Fecha de Finalización";
            dataGridView3.Columns[8].HeaderText = "Estado de la Orden";
            dataGridView3.Columns[7].HeaderText = "Costo";
            dataGridView3.Columns[7].DefaultCellStyle.Format = "c";
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarPredial();
            cargarPredial();
        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
        }

        private void btnPredial_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            txtPredial.Text = openFileDialog1.FileName;
        }

        private void btnSHP_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox2.Text = openFileDialog1.FileName;
        }

        private void btnEscritura_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            txtEscritura.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el predial del ano " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Predial WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Impuesto predial eliminado.");
                        if (File.Exists("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "\\Impuesto" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString()))
                        {
                            File.Delete("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios" + this.Text + "\\Predial\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "\\Impuesto" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
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
                cargarPredial();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string PDF = textBox1.Text;   
            if (!PDF.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis", "CroquisPDF*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(PDF, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(PDF);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis" + "\\CroquisPDF" + ext));
                }
            }
            string SHP = textBox2.Text;
            if (!SHP.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis", "CroquisSHP*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(SHP, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis" + "\\CroquisSHP" + ext));
                }
            }
            string KMZ = textBox4.Text;
            if (!KMZ.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\KMZ");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\KMZ", "KMZ*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(KMZ, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\KMZ");

                    string ext = Path.GetExtension(SHP);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\KMZ" + "\\KMZ" + ext));
                }
            }
            string escritura = txtEscritura.Text;
            if (!escritura.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Escritura");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Escritura", "EscrituraPublica*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(escritura, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Escritura");
                    
                    string ext = Path.GetExtension(escritura);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Escritura" + "\\EscrituraPublica" + ext));
                }
            }
            string certificado = textBox3.Text;
            if (!escritura.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Certificado");
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Certificado", "CertificadoDeTradicionYLibertad*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(escritura, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Certificado");

                    string ext = Path.GetExtension(escritura);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Certificado" + "\\CertificadoDeTradicionYLibertad" + ext));
                }
            }
            MessageBox.Show("Anexos agregados.");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis", "CroquisPDF*");
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
            //if (File.Exists("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\croquis" + "\\croquisPDF"))
            //{
            //    System.Diagnostics.Process.Start("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\croquis" + "\\croquisPDF");
            //}
            //else
            //{
            //    MessageBox.Show("No se encuentra el archivo.", "Error");
            //}
        }
            
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Croquis", "CroquisSHP*");
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
            //if (File.Exists("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\croquis" + "\\croquisSHP"))
            //{
            //    System.Diagnostics.Process.Start("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\croquis" + "\\croquisSHP");
            //}
            //else
            //{
            //    MessageBox.Show("No se encuentra el archivo.", "Error");
            //}
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Escritura");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Escritura", "EscrituraPublica*");
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
            //if (File.Exists("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\escritura" + "\\NumeroDeEscrituraPublica"))
            //{
            //    System.Diagnostics.Process.Start("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\escritura" + "\\NumeroDeEscrituraPublica");
            //}
            //else
            //{
            //    MessageBox.Show("No se encuentra el archivo.", "Error");
            //}
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Predial\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "Impuesto*");            
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
            //if (File.Exists("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\predial\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "\\Impuesto" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString()))
            //{
            //    System.Diagnostics.Process.Start("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\" + this.Text + "\\predial\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "\\Impuesto" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
            //}
            //else
            //{
            //    MessageBox.Show("No se encuentra el archivo.", "Error");
            //}
        }

        private void txtEscritura_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Certificado");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\Certificado", "Certificado*");
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

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox3.Text = openFileDialog1.FileName;
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        public void cargarLotes()
        {
            string query = "SELECT l.ID, l.Codigo, l.Lote FROM BancoTierras AS b INNER JOIN Lotes AS l ON b.ID = l.Predio WHERE b.id = " + id + " UNION ALL SELECT l.ID, l.Codigo, l.Lote FROM BancoTierras AS b INNER JOIN LoteGanadero AS l ON b.ID = l.Predio WHERE b.id = " + id + " UNION ALL SELECT l.ID, l.Codigo, l.Lote FROM BancoTierras AS b INNER JOIN Areas AS l ON b.ID = l.Predio WHERE b.id = " + id;
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
                    dataGridView2.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView2.Rows[i].Cells[1].Value = myReader.GetInt32(1);
                    dataGridView2.Rows[i].Cells[2].Value = myReader.GetString(2);
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

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\KMZ");
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Predios\\" + this.Text + "\\KMZ", "KMZ*");
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

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString(), 1);
            newFrm.Show();
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView3);
        }
    }
}
