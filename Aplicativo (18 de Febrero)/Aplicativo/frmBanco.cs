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
    public partial class frmBanco : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarBanco() {
            string query = "SELECT b.ID, b.Predio, b.Codigo_predio, b.Codigo_catrastal, b.Matricula,b.numeroEscritura, b.Area, p.Nombre, m.Municipio, b.Longitud, b.Latitud, b.Ubicacion FROM Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN BancoTierras AS b ON p.ID = b.Propietario) ON m.ID = b.Municipio WHERE b.Predio <> 'N/A'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView1.DataSource = banco;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].HeaderText = "Codigo Predio";
            dataGridView1.Columns[3].HeaderText = "Codigo Catrastal";
            dataGridView1.Columns[5].HeaderText = "# de Escritura";
            //dataGridView1.Columns.Add("escritura","Escritura Publica");
            //dataGridView1.Columns.Add("predio", "Predio");
            //dataGridView1.Columns.Add("croquis", "Croquis");

            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    dataGridView1.Rows[i].Cells[10].Value = "Escritura";
            //    dataGridView1.Rows[i].Cells[11].Value = "Predio";
            //    dataGridView1.Rows[i].Cells[12].Value = "Croquis";
            //}   
        }

        public void cargarPropietarios() {
            string query = "SELECT * FROM Propietarios";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtPropietario.DataSource = ds.Tables[0];
            txtPropietario.DisplayMember = "Nombre";
            txtPropietario.ValueMember = "ID";
            txtPropietario.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtPropietario.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarMunicipios() {
            string query = "SELECT * FROM Municipio";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtMunicipio.DataSource = ds.Tables[0];
            txtMunicipio.DisplayMember = "Municipio";
            txtMunicipio.ValueMember = "ID";
            txtMunicipio.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtMunicipio.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void buscarBanco()
        {
            string query = "SELECT b.ID, b.Predio, b.Codigo_predio, b.Codigo_catrastal, b.Matricula,b.numeroEscritura, b.Area, p.Nombre, m.Municipio, b.Longitud, b.Latitud FROM Municipio AS m INNER JOIN (Propietarios AS p INNER JOIN BancoTierras AS b ON p.ID = b.Propietario) ON m.ID = b.Municipio ";
            int i = 0;
            if (!txtPredio.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Predio LIKE '%" + txtPredio.Text + "%'";
            }
            if (!txtCodPred.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Codigo_predio LIKE '%" + txtCodPred.Text + "%'";
            }
            if (!txtCodCat.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Codigo_catrastal LIKE '%" + txtCodCat.Text + "%'";
            }
            if (!txtMatricula.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Matricula LIKE '%" + txtMatricula.Text + "%'";
            }
            if (!txtEscritura.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.numeroEscritura LIKE '%" + txtEscritura.Text + "%'";
            }
            if (!txtArea.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Area LIKE '%" + txtArea.Text + "%'";
            }
            if (!txtPropietario.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "p.Nombre LIKE '%" + txtPropietario.Text + "%'";
            }
            if (!txtMunicipio.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Municipio LIKE '%" + txtMunicipio.Text + "%'";
            }
            if (!txtLong.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Longitud LIKE '%" + txtLong.Text + "%'";
            }
            if (!txtLat.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "b.Latitud LIKE '%" + txtLat.Text + "%'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
        }

        public void verificarPrediales() {
            if (DateTime.Now.Month >= 4)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    string query = "SELECT * FROM Predial WHERE Predio = " + dataGridView1.Rows[i].Cells[0].Value.ToString() + " AND ano = " + DateTime.Now.Year;
                    //Ejecutar el query y llenar el GridView.
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    cmd.Connection = conn;
                    conn.Open();
                    OleDbDataReader myReader = cmd.ExecuteReader();
                    int j = 0;
                    try
                    {
                        if (myReader.Read())
                        {
                            j++;
                        }
                    }
                    finally
                    {
                        if (j == 0)
                        {
                            dataGridView1.Rows[i].Cells[2].Style.ForeColor = System.Drawing.Color.Black;
                        }
                        // always call Close when done reading.
                        myReader.Close();
                        // always call Close when done reading.
                        conn.Close();
                    }
                }
            }
        }

        public frmBanco()
        {
            InitializeComponent();
            cargarBanco();
            cargarMunicipios();
            cargarPropietarios();            
            txtMunicipio.SelectedItem = null;
            txtPropietario.SelectedItem = null;           
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            verificarPrediales();
        }

        public void reiniciarTablero()
        {
            txtPredio.Text = "";
            txtCodCat.Text = "";
            txtCodPred.Text = "";
            txtMatricula.Text = "";
            txtEscritura.Text = "";
            txtArea.Text = "";
            txtPropietario.Text = "";
            txtMunicipio.Text = "";
            txtLong.Text = "";
            txtLat.Text = "";
            textBox1.Text = "";
        }

        public void agregarBanco() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO BancoTierras(Predio,Codigo_predio,Codigo_catrastal,Matricula,numeroEscritura,Area,Propietario,Municipio,Longitud,Latitud,Ubicacion) VALUES (@Predio,@Codigo_predio,@Codigo_catrastal,@Matricula,@numeroEscritura,@Area,@Propietario,@Municipio,@Longitud,@Latitud,@Ubicacion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredio.Text;
                cmd.Parameters.Add("@Codigo_predio", OleDbType.VarChar).Value = txtCodPred.Text;
                cmd.Parameters.Add("@Codigo_catrastal", OleDbType.VarChar).Value = txtCodCat.Text;
                cmd.Parameters.Add("@Matricula", OleDbType.VarChar).Value = txtMatricula.Text;
                cmd.Parameters.Add("@numeroEscritura", OleDbType.VarChar).Value = txtEscritura.Text;
                cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtArea.Text;
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = txtPropietario.SelectedValue;
                cmd.Parameters.Add("@Municipio", OleDbType.VarChar).Value = txtMunicipio.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLong.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLat.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Banco de tierras agregado.");
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

        public void modificarBanco()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE BancoTierras SET Predio=@Predio,Codigo_predio=@Codigo_predio,Codigo_catrastal=@Codigo_catrastal,Matricula=@Matricula,numeroEscritura=@numeroEscritura,Area=@Area,Propietario=@Propietario,Municipio=@Municipio,Longitud=@Longitud,Latitud=@Latitud,Ubicacion=@Ubicacion WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Predio", OleDbType.VarChar).Value = txtPredio.Text;
                cmd.Parameters.Add("@Codigo_predio", OleDbType.VarChar).Value = txtCodPred.Text;
                cmd.Parameters.Add("@Codigo_catrastal", OleDbType.VarChar).Value = txtCodCat.Text;
                cmd.Parameters.Add("@Matricula", OleDbType.VarChar).Value = txtMatricula.Text;
                cmd.Parameters.Add("@numeroEscritura", OleDbType.VarChar).Value = txtEscritura.Text;
                cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = txtArea.Text;
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = txtPropietario.SelectedValue;
                cmd.Parameters.Add("@Municipio", OleDbType.VarChar).Value = txtMunicipio.SelectedValue;
                cmd.Parameters.Add("@Longitud", OleDbType.VarChar).Value = txtLong.Text;
                cmd.Parameters.Add("@Latitud", OleDbType.VarChar).Value = txtLat.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Banco de tierras modficiado.");
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

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!txtPredio.Text.Equals(""))
            {
                if (!txtMunicipio.Text.Equals(""))
                {
                    if (!txtPropietario.Text.Equals(""))
                    {
                        agregarBanco();
                        cargarBanco();
                        reiniciarTablero();
                    }
                    else
                        MessageBox.Show("Favor seleccionar propietario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                    MessageBox.Show("Favor seleccionar Municipio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
                MessageBox.Show("Favor ingresar nombre del predio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmInfoBanco newFrm = new frmInfoBanco(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            txtPredio.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtCodCat.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtCodPred.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtMatricula.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtEscritura.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtArea.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtPropietario.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtMunicipio.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtLong.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            txtLat.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            if (!txtPredio.Text.Equals(""))
            {
                if (!txtMunicipio.Text.Equals(""))
                {
                    if (!txtPropietario.Text.Equals(""))
                    {
                        modificarBanco();
                        cargarBanco();
                        reiniciarTablero();
                    }
                    else
                        MessageBox.Show("Favor seleccionar propietario.", "Error");
                }
                else
                    MessageBox.Show("Favor seleccionar Municipio.", "Error");
            }
            else
                MessageBox.Show("Favor ingresar nombre del predio.", "Error");
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar a " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM BancoTierras WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Banco de Tierra eliminado.");
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
                cargarBanco();
                reiniciarTablero();
            }
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            buscarBanco();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmLotes newFrm = new frmLotes();
            newFrm.Show();
        }

        private void btnPropietarios_Click(object sender, EventArgs e)
        {
            frmPropietarios newFrm = new frmPropietarios();
            newFrm.Show();
        }

        private void btnMunicipio_Click(object sender, EventArgs e)
        {
            frmMunicipios newFrm = new frmMunicipios();
            newFrm.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
