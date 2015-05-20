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
    public partial class frmCuchillasOrden : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int tipousuario = 1;
        int orden = 0;
        List<int> Cuchilla = new List<int>();
        List<int> Maquina = new List<int>();
        List<string> Puesto = new List<string>();        

        public frmCuchillasOrden(int op, int tipo)
        {
            InitializeComponent();
            tipousuario = tipo;
            orden = op;
            cargarFormatoEntrega(dataGridView2, "", "");
            cargarCuchillas();
            cargarCuchillasAsignadas(orden, dataGridView2);
            Variables.cargar(dataGridView1, "SELECT c.Cuchilla, Cuchillas.Codigo, (m.Placa + ' / ' + m.Marca + ' / ' + m.Modelo) As Maquinaria, c.Puesto FROM (cuchillaAsignadas AS c INNER JOIN Maquinarias AS m ON c.Maquina = m.ID) INNER JOIN Cuchillas ON c.Cuchilla = Cuchillas.Id WHERE OP = " + op);            
        }

        public void cargarFormatoEntrega(DataGridView data, string query, string display)
        {
            DataGridViewCheckBoxColumn check = new DataGridViewCheckBoxColumn();
            check.HeaderText = "";
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            //combo.HeaderText = "Cuchilla";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add(check);
            data.Columns.Add("Column3", "Cuchilla");
            //cargarDetalle(combo, "SELECT * FROM Cuchillas", "Codigo", data);
            //data.Columns.Add(combo);
            combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Maquina";
            cargarDetalle(combo, "SELECT ID, (Placa + ' / ' + Marca + ' / ' + Modelo) As Maquina FROM Maquinarias WHERE Tipo = 'Aserradero'", "Maquina", data);
            data.Columns.Add(combo);
            data.Columns.Add("Column4", "Puesto");
            //combo = new DataGridViewComboBoxColumn();
            //data.Columns.Add(combo);
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 40;            
            data.Columns[5].FillWeight = 40;               
        }

        public void cargarDetalle(DataGridViewComboBoxColumn combo, string query, string display, DataGridView data)
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
        }

        public void cargarCuchillas()
        {
            string query = "SELECT ID,Codigo,Maquina,puestoMaquina FROM Cuchillas";
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
                    dataGridView2.Rows[i].Cells[0].Value = i + 1;
                    dataGridView2.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView2.Rows[i].Cells[3].Value = myReader.GetString(1);
                    dataGridView2.Rows[i].Cells[4].Value = myReader.GetInt32(2);
                    dataGridView2.Rows[i].Cells[5].Value = myReader.GetString(3);
                    Cuchilla.Add(myReader.GetInt32(0));
                    Maquina.Add(myReader.GetInt32(2));
                    Puesto.Add(myReader.GetString(3));
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

        public void cargarCuchillasAsignadas(int op, DataGridView data)
        {
            string query = "SELECT ID,Cuchilla,Maquina,Puesto FROM cuchillaAsignadas WHERE OP = " + op;
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
                    for (int j = 0; j < data.Rows.Count; j++)
                    {
                        if (data.Rows[j].Cells[1].Value.ToString().Equals(myReader.GetInt32(1).ToString()))
                        {
                            data.Rows[j].Cells[2].Value = true;
                            data.Rows[j].Cells[4].Value = myReader.GetInt32(2);
                            data.Rows[j].Cells[5].Value = myReader.GetString(3);
                        }
                    }
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
            if (existeAsignacion())
                eliminarAsignacion();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dataGridView2.Rows[i].Cells[2].Value) == true)
                {
                    if (Int32.Parse(dataGridView2.Rows[i].Cells[4].Value.ToString()) != Maquina[i] || !dataGridView2.Rows[i].Cells[5].Value.ToString().Equals(Puesto[i]))
                        agregarMovimiento(dataGridView2, i);
                    agregarAsignacion(i, dataGridView2);
                }
            }
            MessageBox.Show("Cuchillas asignadas.");
        }

        public void agregarMovimiento(DataGridView data, int i)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            cmd = new OleDbCommand("INSERT INTO cuchillaMaquina (Cuchilla,Maquina,Fecha,Puesto) VALUES (@Cuchilla,@Maquina,@Fecha,@Puesto)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value;
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
                cmd.Parameters.Add("@Puesto", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value;
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

        public void agregarAsignacion(int i, DataGridView data)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO cuchillaAsignadas (Cuchilla,Maquina,Puesto,OP) VALUES (@Cuchilla,@Maquina,@Puesto,@OP)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value;
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value;
                cmd.Parameters.Add("@Puesto", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value;
                cmd.Parameters.Add("@OP", OleDbType.VarChar).Value = orden;
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

        public void eliminarAsignacion()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("DELETE FROM cuchillaAsignadas WHERE OP = " + orden);
            cmd.Connection = conn;
            conn.Open();

            if (conn.State == ConnectionState.Open)
            {
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

        public bool existeAsignacion()
        {
            string query = "SELECT * FROM cuchillaAsignadas WHERE OP = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())                
                    return true;               
                else
                    return false;
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        public void generarEvento(string evento, int cuchilla)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            cmd = new OleDbCommand("INSERT INTO cuchillaEvento (Cuchilla,Evento,Fecha,Hora,OP) VALUES (@Cuchilla,@Evento,@Fecha,@Hora,@OP)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = cuchilla;
                cmd.Parameters.Add("@Evento", OleDbType.VarChar).Value = evento;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year ;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = DateTime.Now.ToString("HH") + ":" + DateTime.Now.ToString("mm") + ":" + DateTime.Now.ToString("ss");
                cmd.Parameters.Add("@OP", OleDbType.VarChar).Value = orden;
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

        public void registrarHora(string evento, int cuchilla)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            cmd = new OleDbCommand("INSERT INTO cuchillaHoras (Cuchilla,Hora,OP,Fecha) VALUES (@Cuchilla,@Hora,@OP,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = cuchilla;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = evento;
                cmd.Parameters.Add("@OP", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Horas de uso registrada.");
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


        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            generarEvento("Afilada", Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            Variables.cargar(dataGridView4, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
            dataGridView4.Columns[1].HeaderText = "Afilado/Triscado";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Variables.cargar(dataGridView4, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
            dataGridView4.Columns[1].HeaderText = "Afilado/Triscado";
            Variables.cargar(dataGridView3, "SELECT ID,Hora,Fecha FROM cuchillaHoras WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            generarEvento("Triscada", Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            Variables.cargar(dataGridView4, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
            dataGridView4.Columns[1].HeaderText = "Afilado/Triscado";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            generarEvento("Afilada", Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            generarEvento("Triscada", Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            Variables.cargar(dataGridView4, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
            dataGridView4.Columns[1].HeaderText = "Afilado/Triscado";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la " + dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                string id = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM cuchillaEvento WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[1].Value.ToString() + " eliminada.");
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
                Variables.cargar(dataGridView4, "SELECT ID,Evento,Fecha,Hora FROM cuchillaEvento WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
                dataGridView4.Columns[1].HeaderText = "Afilado/Triscado";

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            registrarHora(textBox1.Text, Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            Variables.cargar(dataGridView3, "SELECT ID,Hora,Fecha FROM cuchillaHoras WHERE Cuchilla = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString() + " AND OP = " + orden);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Variables.cargar(dataGridView1, "SELECT c.Cuchilla, Cuchillas.Codigo, (m.Placa + ' / ' + m.Marca + ' / ' + m.Modelo) As Maquinaria, c.Puesto FROM (cuchillaAsignadas AS c INNER JOIN Maquinarias AS m ON c.Maquina = m.ID) INNER JOIN Cuchillas ON c.Cuchilla = Cuchillas.Id WHERE OP = " + orden);
       }

    }
}
