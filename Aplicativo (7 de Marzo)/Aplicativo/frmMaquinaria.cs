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
    public partial class frmMaquinaria : Form
    {
        //Base de datos.
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarMaquinaria()
        {
            string query = "SELECT m.ID, m.Placa, m.Tipo, m.Marca, m.Modelo, e.Estado, m.Ano, m.Ano_Fabricacion,m.Tipo_Combustible, m.Descripcion, p.Propietario, m.aceiteMotor, m.aceiteHidraulico, m.aceiteCajaV, m.aceiteDiferencial,m.Ubicacion,m.Horometro,m.numPuestos,m.valorHora FROM propietariosMaquina AS p INNER JOIN (Estados AS e INNER JOIN Maquinarias AS m ON e.ID = m.Estado) ON p.ID = m.Propietario;";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinarias = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(maquinarias);
            dataGridView1.DataSource = maquinarias;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[6].HeaderText = "Año de Compra";
            dataGridView1.Columns[7].HeaderText = "Año de Fabricación";
            dataGridView1.Columns[8].HeaderText = "Tipo de Combustible";
            dataGridView1.Columns[11].HeaderText = "Ref. Aceite Motor";
            dataGridView1.Columns[12].HeaderText = "Ref. Aceite Hidraulico";
            dataGridView1.Columns[13].HeaderText = "Ref. Aceite de Caja V";
            dataGridView1.Columns[14].HeaderText = "Ref. Aceite de Diferenciales";
            dataGridView1.Columns[15].HeaderText = "Ubicación Fisica";
            dataGridView1.Columns[17].HeaderText = "# Puestos";
            dataGridView1.Columns[18].HeaderText = "Valor/Hora";
            dataGridView1.Columns[18].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[11].Visible = false;
            dataGridView1.Columns[12].Visible = false;
            dataGridView1.Columns[13].Visible = false;
            dataGridView1.Columns[14].Visible = false;
            dataGridView1.Columns[15].Visible = false;
            dataGridView1.Columns[17].Visible = false;         
        }

        public void buscarMaquinaria()
        {
            string query = "SELECT m.ID, m.Placa, m.Tipo, m.Marca, m.Modelo, e.Estado, m.Ano, m.Ano_Fabricacion,m.Tipo_Combustible, m.Descripcion, p.Propietario, m.aceiteMotor, m.aceiteHidraulico, m.aceiteCajaV, m.aceiteDiferencial,m.Ubicacion,m.Horometro FROM propietariosMaquina AS p INNER JOIN (Estados AS e INNER JOIN Maquinarias AS m ON e.ID = m.Estado) ON p.ID = m.Propietario ";
            int i = 0;
            if (!txtPlaca.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Placa LIKE '%" + txtPlaca.Text + "%'";
            }
            if (!txtTipo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Tipo LIKE '%" + txtTipo.Text + "%'";
            }
            if (!txtMarca.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Marca LIKE '%" + txtMarca.Text + "%'";
            }
            if (!txtModelo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Modelo LIKE '%" + txtModelo.Text + "%'";
            }
            if (!txtCombustible.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;                
                query += "m.Tipo_Combustible LIKE '%" + txtCombustible.Text + "%'";
            }
            if (!txtFabricacion.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Ano_Fabricacion LIKE '%" + txtFabricacion.Text + "%'";
            }
            if (!txtAno.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Ano LIKE '%" + txtAno.Text + "%'";
            }
            if (!txtPropietario.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "p.Propietario LIKE '%" + txtPropietario.Text + "%'";
            } 
            if (!txtDescripcion.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Descripcion LIKE '%" + txtDescripcion.Text + "%'";
            }
            if (!txtAceite4.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.aceiteDiferencial LIKE '%" + txtAceite4.Text + "%'";
            }
            if (!txtAceite1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.aceiteMotor LIKE '%" + txtAceite1.Text + "%'";
            }
            if (!txtAceite2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.aceiteHidraulico LIKE '%" + txtAceite2.Text + "%'";
            }
            if (!txtAceite3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.aceiteCajaV LIKE '%" + txtAceite3.Text + "%'";
            }
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Ubicacion LIKE '%" + textBox1.Text + "%'";
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "m.Horometro LIKE '%" + textBox2.Text + "%'";
            }
            //if (!txtEncargado.Text.Equals(""))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "s.Supervisor LIKE '%" + txtEncargado.Text + "%'";
            //}
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
        }

        public void cargarUnidades()
        {
            string query = "SELECT * FROM propietariosMaquina";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtPropietario.DataSource = ds.Tables[0];
            txtPropietario.DisplayMember = "Propietario";
            txtPropietario.ValueMember = "ID";
            txtPropietario.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtPropietario.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarEstados()
        {
            string query = "SELECT * FROM Estados WHERE Tipo = 'Maquinarias'";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtEstado.DataSource = ds.Tables[0];
            txtEstado.DisplayMember = "Estado";
            txtEstado.ValueMember = "ID";
            txtEstado.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtEstado.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void agregarMaquinaria()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Maquinarias (Placa,Tipo,Marca,Ano,Modelo,Estado,Ano_Fabricacion,Tipo_Combustible,Descripcion,Propietario,aceiteMotor,aceiteHidraulico,aceiteCajaV,aceiteDiferencial,Ubicacion,Horometro,numPuestos, valorHora) VALUES (@Placa,@Tipo,@Marca,@Ano,@Modelo,@Estado,@Ano_Fabricacion,@Tipo_Combustible,@Descripcion,@Propietario,@aceiteMotor,@aceiteHidraulico,@aceiteCajaV,@aceiteDiferencial,@Ubicacion,@Horometro,@numPuestos, @valorHora)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = txtPlaca.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                cmd.Parameters.Add("@Marca", OleDbType.VarChar).Value = txtMarca.Text;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = txtAno.Text;
                cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = txtModelo.Text;
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = txtEstado.SelectedValue;
                cmd.Parameters.Add("@Ano_Fabricacion", OleDbType.VarChar).Value = txtFabricacion.Text;
                cmd.Parameters.Add("@Tipo_Combustible", OleDbType.VarChar).Value = txtCombustible.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;                
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = txtPropietario.SelectedValue;
                cmd.Parameters.Add("@aceiteMotor", OleDbType.VarChar).Value = txtAceite1.Text;
                cmd.Parameters.Add("@aceiteHidraulico", OleDbType.VarChar).Value = txtAceite2.Text;
                cmd.Parameters.Add("@aceiteCajaV", OleDbType.VarChar).Value = txtAceite3.Text;
                cmd.Parameters.Add("@aceiteDiferencial", OleDbType.VarChar).Value = txtAceite4.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;
                if(textBox2.Text.Equals(""))
                    cmd.Parameters.Add("@Horometro", OleDbType.VarChar).Value = textBox2.Text;
                else
                    cmd.Parameters.Add("@Horometro", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@numPuestos", OleDbType.VarChar).Value = txtNumPuestos.Text;
                cmd.Parameters.Add("@valorHora", OleDbType.VarChar).Value = txtValor.Text;
                //cmd.Parameters.Add("@Encargado", OleDbType.VarChar).Value = txtEncargado.SelectedValue;

                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Maquina agregada.");
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

        public void eliminarMaquinaria()
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la maquina " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() + " / " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString() + " / " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {

                    string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("DELETE FROM Maquinarias WHERE id = " + id);
                    cmd.Connection = conn;
                    conn.Open();

                    if (conn.State == ConnectionState.Open)
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Maquinaria eliminada.");
                            conn.Close();
                            cargarMaquinaria();
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
            else
            {
                MessageBox.Show("Favor seleccionar la placa de la maquina.", "Error");
            }
        }

        //public void cargarSupervisores()
        //{
        //    string query = "SELECT id, (Nombres + ' ' + Apellidos) As Nombre FROM Trabajadores WHERE Supervisor = 'Si'";
        //    //Ejecutar el query y llenar el GridView.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    DataTable departamentos = new DataTable();
        //    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        //    DataSet ds = new DataSet();
        //    da.Fill(ds);
        //    txtEncargado.DataSource = ds.Tables[0];
        //    txtEncargado.DisplayMember = "Nombre";
        //    txtEncargado.ValueMember = "ID";
        //    txtEncargado.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        //    txtEncargado.AutoCompleteSource = AutoCompleteSource.ListItems;
        //}

        //public void cargarOperadores()
        //{
        //    string query = "SELECT id, (Nombres + ' ' + Apellidos) As Nombre FROM Trabajadores WHERE Supervisor = 'No'";
        //    //Ejecutar el query y llenar el GridView.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    DataTable departamentos = new DataTable();
        //    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        //    DataSet ds = new DataSet();
        //    da.Fill(ds);
        //    txtOperador.DataSource = ds.Tables[0];
        //    txtOperador.DisplayMember = "Nombre";
        //    txtOperador.ValueMember = "ID";
        //    txtOperador.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        //    txtOperador.AutoCompleteSource = AutoCompleteSource.ListItems;
        //}

        public void reiniciarTablero() {
            txtPlaca.Text = "";
            txtTipo.Text = "";
            txtMarca.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            txtModelo.Text = "";
            txtAno.Text = "";
            txtEstado.Text = "";
            txtFabricacion.Text = "";
            txtCombustible.Text = "";
            txtDescripcion.Text = "";
            txtAceite1.Text = "";
            txtAceite2.Text = "";
            txtAceite3.Text = "";
            txtAceite4.Text = "";
            txtPropietario.Text = "";
            txtNumPuestos.Text = "0";
        }

        public frmMaquinaria()
        {
            InitializeComponent();            
            cargarMaquinaria();
            cargarUnidades();
            cargarEstados();
            //cargarSupervisores();
            //cargarOperadores();
            //txtOperador.SelectedItem = null;
            //txtEncargado.SelectedItem = null;
            txtPropietario.SelectedItem = null;
            txtEstado.SelectedItem = null;
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font,FontStyle.Underline);
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            eliminarMaquinaria();
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            frmVerMaquinaria newFrm = new frmVerMaquinaria(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            newFrm.Show();
        }

        public bool existePlaca(string codigo)
        {
            string query = "SELECT ID FROM Maquinarias WHERE Placa = '" + codigo + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            bool existe = true;
            try
            {
                if (myReader.HasRows)
                {
                    existe = true;
                }
                else
                {
                    existe = false;
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return existe;
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!existePlaca(txtPlaca.Text))
            {
                agregarMaquinaria();
                cargarMaquinaria();
            }
            else
                MessageBox.Show("Placa ya existe.");
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            buscarMaquinaria();
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmVerMaquinaria newFrm = new frmVerMaquinaria(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            txtPlaca.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtTipo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtMarca.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtModelo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtAno.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtEstado.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtFabricacion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtCombustible.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtDescripcion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();            
            txtAceite1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            txtAceite2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            txtAceite3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            txtAceite4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            txtNumPuestos.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            txtValor.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            txtPropietario.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
        }

        public void modificarMaquinaria() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Maquinarias SET Placa=@Placa,Tipo=@Tipo,Marca=@Marca,Modelo=@Modelo,Estado=@Estado,Ano=@Ano,Ano_Fabricacion=@Ano_Fabricacion,Tipo_Combustible=@Tipo_Combustible,Descripcion=@Descripcion,Propietario=@Propietario,aceiteMotor=@aceiteMotor,aceiteHidraulico=@aceiteHidraulico,aceiteCajaV=@aceiteCajaV,aceiteDiferencial=@aceiteDiferencial,Ubicacion=@Ubicacion,Horometro=@Horometro,numPuestos=@numPuestos,valorHora=@valorHora WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = txtPlaca.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                cmd.Parameters.Add("@Marca", OleDbType.VarChar).Value = txtMarca.Text;
                cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = txtModelo.Text;
                cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = txtEstado.SelectedValue;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = txtAno.Text;
                cmd.Parameters.Add("@Ano_Fabricacion", OleDbType.VarChar).Value = txtFabricacion.Text;
                cmd.Parameters.Add("@Tipo_Combustible", OleDbType.VarChar).Value = txtCombustible.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = txtPropietario.SelectedValue;
                cmd.Parameters.Add("@aceiteMotor", OleDbType.VarChar).Value = txtAceite1.Text;
                cmd.Parameters.Add("@aceiteHidraulico", OleDbType.VarChar).Value = txtAceite2.Text;
                cmd.Parameters.Add("@aceiteCajaV", OleDbType.VarChar).Value = txtAceite3.Text;
                cmd.Parameters.Add("@aceiteDiferencial", OleDbType.VarChar).Value = txtAceite4.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;
                if (textBox2.Text.Equals(""))
                    cmd.Parameters.Add("@Horometro", OleDbType.VarChar).Value = textBox2.Text;
                else
                    cmd.Parameters.Add("@Horometro", OleDbType.VarChar).Value = textBox2.Text; cmd.Parameters.Add("@numPuestos", OleDbType.VarChar).Value = txtNumPuestos.Text;
                cmd.Parameters.Add("@valorHora", OleDbType.VarChar).Value = txtValor.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Maquina modificada.");
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

        private void btnModificar_Click(object sender, EventArgs e)
        {
            modificarMaquinaria();
            cargarMaquinaria();
            reiniciarTablero();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmMaquinaPropietarios newFrm = new frmMaquinaPropietarios();
            newFrm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmMaquinaEstado newFrm = new frmMaquinaEstado("Maquinarias");
            newFrm.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                dataGridView1.Cursor = Cursors.Hand;
            }
            else
            {
                dataGridView1.Cursor = Cursors.Default;
            }
        }
    }
}
