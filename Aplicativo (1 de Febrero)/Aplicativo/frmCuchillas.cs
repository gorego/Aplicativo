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
    public partial class frmCuchillas : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmCuchillas()
        {
            InitializeComponent();
            Variables.cargar(comboBox1, "SELECT ID, (Placa + ' / ' + Marca + ' / ' + Modelo) As Maquina  FROM Maquinarias WHERE Tipo = 'Asserin';", "Maquina");
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            cargar();
            comboBox1.SelectedItem = null;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        public void cargar()
        {
            Variables.cargar(dataGridView1, "SELECT c.Id, c.Codigo, (m.Placa + ' / ' + m.Marca + ' / ' + m.Modelo), c.puestoMaquina, c.Marca, c.Referencia, c.Ancho, c.Grueso, c.Largo, c.Valor, c.fechaIngreso, c.Estado FROM Maquinarias AS m INNER JOIN Cuchillas AS c ON m.ID = c.Maquina;");
            dataGridView1.Columns[2].HeaderText = "Maquina";
            dataGridView1.Columns[2].HeaderText = "Puesto en Maquina";
            dataGridView1.Columns[2].HeaderText = "Fecha Ingreso";
        }

        public void modificarCuchilla(int tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            if(tipo == 0)
                cmd = new OleDbCommand("INSERT INTO Cuchillas (Maquina,puestoMaquina,Marca,Referencia,Ancho,Grueso,Largo,Valor,fechaIngreso,Estado,Codigo) VALUES (@Maquina,@puestoMaquina,@Marca,@Referencia,@Ancho,@Grueso,@Largo,@Valor,@fechaIngreso,@Estado,@Codigo)");
            else
                cmd = new OleDbCommand("UPDATE Cuchillas SET Maquina=@Maquina, puestoMaquina=@puestoMaquina, Marca=@Marca, Referencia=@Referencia, Ancho=@Ancho, Grueso=@Grueso, Largo=@Largo, Valor=@Valor, fechaIngreso=@fechaIngreso, Estado=@Estado, Codigo=@Codigo WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@puestoMaquina", OleDbType.VarChar).Value = txtPuesto.Text;
                cmd.Parameters.Add("@Marca", OleDbType.VarChar).Value = txtMarca.Text;
                cmd.Parameters.Add("@Referencia", OleDbType.VarChar).Value = txtReferencia.Text;
                cmd.Parameters.Add("@Ancho", OleDbType.VarChar).Value = txtAncho.Text;
                cmd.Parameters.Add("@Grueso", OleDbType.VarChar).Value = txtGrueso.Text;
                cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = txtLargo.Text;
                cmd.Parameters.Add("@Valor", OleDbType.VarChar).Value = txtValor.Text;
                cmd.Parameters.Add("@fechaIngreso", OleDbType.VarChar).Value = dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = comboBox2.Text;
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = txtCodigo.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    if (tipo == 0)
                        MessageBox.Show("Cuchilla agregada.");
                    else
                        MessageBox.Show("Cuchilla modificada."); conn.Close();
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

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            modificarCuchilla(0);
            cargar();
            reiniciar();
            int id = getMaxID();
            agregarMaquina(0, id.ToString());
        }

        public void agregarMaquina(int tipo, string id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            if (tipo == 0)
                cmd = new OleDbCommand("INSERT INTO cuchillaMaquina (Cuchilla,Maquina,Fecha,Puesto) VALUES (@Cuchilla,@Maquina,@Fecha,@Puesto)");
            else
                cmd = new OleDbCommand("UPDATE cuchillaMaquina SET Cuchilla=@Cuchilla, Maquina=@Maquina, Fecha=@Fecha WHERE Cuchilla = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = id;
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = txtPuesto.Text;
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

        public int getMaxID()
        {
            string query = "SELECT MAX(id) FROM CuchillaS";
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

        public void reiniciar()
        {
            comboBox1.Text = "";
            txtPuesto.Text = "";
            txtMarca.Text = "";
            txtReferencia.Text = "";
            txtAncho.Text = "0";
            txtGrueso.Text = "0";
            txtLargo.Text = "0";
            txtValor.Text = "0";
            comboBox2.Text = "";
            txtCodigo.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            modificarCuchilla(1);
            cargar();
            reiniciar();
            agregarMaquina(0, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la cuchilla " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {                
                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Cuchillas WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Cuchilla eliminada.");
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
                cargar();
                reiniciar();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            comboBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtPuesto.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtMarca.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtReferencia.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtAncho.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtGrueso.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtLargo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtValor.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            comboBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            txtCodigo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmInfoCuchilla newFrm = new frmInfoCuchilla(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            newFrm.Show();
        }
    }
}
