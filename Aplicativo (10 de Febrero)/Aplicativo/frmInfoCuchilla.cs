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
    public partial class frmInfoCuchilla : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string cuchilla = "";

        public frmInfoCuchilla(string id)
        {
            InitializeComponent();
            Variables.cargar(dataGridView1, "SELECT (m.Placa + ' / ' + m.Marca + ' / ' + m.Modelo) As Maquina, c.Puesto, c.Fecha FROM Maquinarias AS m INNER JOIN cuchillaMaquina AS c ON m.ID = c.Maquina WHERE Cuchilla = " + id);
            dataGridView1.Columns[0].Visible = true;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm"; // Only use hours and minutes
            dateTimePicker2.ShowUpDown = true;
            cuchilla = id;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            Variables.cargar(dataGridView2, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + id);
            dataGridView2.Columns[1].HeaderText = "Afilado/Triscado";
            getTotal(dataGridView2);
        }

        public void getTotal(DataGridView data)
        {
            int total = 0, total2 = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (data.Rows[i].Cells[1].Value.ToString().Equals("Afilada"))
                    total++;
                else
                    total2++;
            }
            label2.Text = "# de Afiladas : " + total + "     # de Triscadas : " + total2;
        }

        public void generarEvento(string evento)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
                cmd = new OleDbCommand("INSERT INTO cuchillaEvento (Cuchilla,Evento,Fecha,Hora) VALUES (@Cuchilla,@Evento,@Fecha,@Hora)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = cuchilla;
                cmd.Parameters.Add("@Evento", OleDbType.VarChar).Value = evento;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = dateTimePicker2.Value.ToString("HH") + ":" + dateTimePicker2.Value.ToString("mm");
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

        private void button1_Click(object sender, EventArgs e)
        {
            generarEvento("Afilada");
            Variables.cargar(dataGridView2, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + cuchilla);
            dataGridView2.Columns[1].HeaderText = "Afilado/Triscado";
            getTotal(dataGridView2);            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            generarEvento("Triscada");
            Variables.cargar(dataGridView2, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + cuchilla);
            dataGridView2.Columns[1].HeaderText = "Afilado/Triscado";
            getTotal(dataGridView2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            generarEvento("Afilada");
            generarEvento("Triscada");
            Variables.cargar(dataGridView2, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + cuchilla);
            dataGridView2.Columns[1].HeaderText = "Afilado/Triscado";
            getTotal(dataGridView2);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                string id = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM cuchillaEvento WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString() + " eliminada.");
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
                Variables.cargar(dataGridView2, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + cuchilla);
                dataGridView2.Columns[1].HeaderText = "Afilado/Triscado";
                getTotal(dataGridView2);
            }
        }
    }
}
