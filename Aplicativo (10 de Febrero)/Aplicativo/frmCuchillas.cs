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
        int tipousuario = 1;

        public frmCuchillas(int tipo)
        {
            InitializeComponent();
            Variables.cargar(comboBox1, "SELECT ID, (Placa + ' / ' + Marca + ' / ' + Modelo) As Maquina  FROM Maquinarias WHERE Tipo = 'Asserin';", "Maquina");
            Variables.cargar(txtCuchilla, "SELECT ID, (Codigo + ' / ' + Marca + ' / ' + Modelo) As Cuchilla  FROM Insumos WHERE Clase = 'Cuchilla';", "Cuchilla");
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            tipousuario = tipo;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            cargar();
            comboBox1.SelectedItem = null;
            txtCuchilla.SelectedItem = null;
            if (tipousuario != 0)
            {
                txtValor.Visible = false;
                label5.Visible = false;
                btnEliminar.Enabled = false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        public void cargar()
        {
            Variables.cargar(dataGridView1, "SELECT c.Id, c.Codigo, (m.Placa + ' / ' + m.Marca + ' / ' + m.Modelo) , c.puestoMaquina, i.Marca, i.Modelo, i.Ancho, i.Grueso, i.Largo, i.Costo_Unitario, c.fechaIngreso, c.Estado, i.ID FROM (Cuchillas AS c INNER JOIN Maquinarias AS m ON c.Maquina = m.ID) INNER JOIN Insumos AS i ON c.Insumo = i.ID;");
            dataGridView1.Columns[2].HeaderText = "Maquina";
            dataGridView1.Columns[3].HeaderText = "Puesto en Maquina";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[9].HeaderText = "Valor";
            dataGridView1.Columns[10].HeaderText = "Fecha Ingreso";
            dataGridView1.Columns[12].Visible = false;
            if(tipousuario != 0)
                dataGridView1.Columns[9].Visible = false;
        }

        public void modificarCuchilla(int tipo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            if(tipo == 0)
                cmd = new OleDbCommand("INSERT INTO Cuchillas (Maquina,puestoMaquina,Insumo,fechaIngreso,Estado,Codigo) VALUES (@Maquina,@puestoMaquina,@Insumo,@fechaIngreso,@Estado,@Codigo)");
            else
                cmd = new OleDbCommand("UPDATE Cuchillas SET Maquina=@Maquina, puestoMaquina=@puestoMaquina, Insumo=@Insumo, fechaIngreso=@fechaIngreso, Estado=@Estado, Codigo=@Codigo WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@puestoMaquina", OleDbType.VarChar).Value = txtPuesto.Text;
                cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = txtCuchilla.SelectedValue;
                cmd.Parameters.Add("@fechaIngreso", OleDbType.VarChar).Value = dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = comboBox2.Text;
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
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
            int id = getMaxID();
            agregarMaquina(0, id.ToString());
            reiniciar();
        }

        public void agregarMaquina(int tipo, string id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand();
            if (tipo == 0)
                cmd = new OleDbCommand("INSERT INTO cuchillaMaquina (Cuchilla,Maquina,Fecha,Puesto) VALUES (@Cuchilla,@Maquina,@Fecha,@Puesto)");
            else
                cmd = new OleDbCommand("UPDATE cuchillaMaquina SET Cuchilla=@Cuchilla, Maquina=@Maquina,Fecha=@Fecha, Puesto=@Puesto WHERE Cuchilla = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cuchilla", OleDbType.VarChar).Value = id;
                cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
                cmd.Parameters.Add("@Puesto", OleDbType.VarChar).Value = txtPuesto.Text;
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
            agregarMaquina(0, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            reiniciar();
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
            txtValor.Text = String.Format("{0:c}",dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString());
            dateTimePicker1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            comboBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            txtCodigo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtCuchilla.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox1.Text = getUso(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            textBox2.Text = getEvento(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(),"Triscada").ToString();
            textBox3.Text = getEvento(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(),"Afilada").ToString();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmInfoCuchilla newFrm = new frmInfoCuchilla(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            newFrm.Show();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public int getNumPuestos(string id)
        {
            string query = "SELECT numPuestos FROM Maquinarias WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int puestos = 0;
            try
            {
                while (myReader.Read())
                {
                    puestos = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return puestos;
        }

        public string getUso(string id)
        {
            string query = "SELECT Sum(Hora) FROM cuchillaHoras WHERE Cuchilla = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string puestos = "0";
            try
            {
                while (myReader.Read())
                {
                    puestos = myReader.GetValue(0).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return puestos;
        }

        public int getEvento(string id, string evento)
        {
            string query = "SELECT COUNT(Evento) FROM cuchillaEvento WHERE Cuchilla = " + id + " AND Evento = '" + evento + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int puestos = 0;
            try
            {
                while (myReader.Read())
                {
                    puestos = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return puestos;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!comboBox1.Text.Equals("") && !comboBox1.Text.Equals("System.Data.DataRowView"))
            {
                txtPuesto.Items.Clear();
                int puestos = getNumPuestos(comboBox1.SelectedValue.ToString());
                for (int i = 0; i < puestos; i++)
                {
                    txtPuesto.Items.Add(i+1);
                }
                if (puestos == 0)
                {
                    txtPuesto.Items.Add(0);
                }
                string[] maquina = comboBox1.Text.Split('/');
                string placa = maquina[0].Trim();                
                string puesto = "0";
                if (!txtPuesto.Text.Equals(""))
                    puesto = txtPuesto.Text;
                txtCodigo.Text = "CH - " + placa + " - " + puesto.PadLeft(4, '0');

            }
        }

        public void getInsumo(string id)
        {
            string query = "SELECT Marca,Modelo,Ancho,Grueso,Largo, Costo_Unitario FROM Insumos WHERE ID = " + id;
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
                    txtMarca.Text = myReader.GetValue(0).ToString();
                    txtReferencia.Text = myReader.GetValue(1).ToString();
                    txtAncho.Text = myReader.GetDouble(2).ToString();
                    txtGrueso.Text = myReader.GetValue(3).ToString();
                    txtLargo.Text = myReader.GetValue(4).ToString();
                    txtValor.Text = String.Format("{0:c}",myReader.GetValue(5).ToString());
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

        private void txtPuesto_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!comboBox1.Text.Equals("") && !txtPuesto.Text.Equals(""))
            {
                string puesto = "0";
                if (!txtPuesto.Text.Equals(""))
                    puesto = txtPuesto.Text;
                string[] maquina = comboBox1.Text.Split('/');
                string placa = maquina[0].Trim();   
                txtCodigo.Text = "CH - " + placa + " - " + puesto.PadLeft(4, '0');
            }
        }

        private void txtCuchilla_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!txtCuchilla.Text.Equals("") && !txtCuchilla.Text.Equals("System.Data.DataRowView"))
            {
                getInsumo(txtCuchilla.SelectedValue.ToString());
            }
        }
    }
}
