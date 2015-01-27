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
    public partial class frmInsumos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        List<string> maquinarias = new List<string>();
        int tipousuario;

        public void cargarInsumos()
        {
            string query = "SELECT i.ID, i.Codigo, i.Clase, i.Marca, i.Modelo, i.Serial, i.Descripcion, p.Proveedor, i.Unidad_Medida, i.Costo_Unitario, i.Cantidad_Stock, i.Stock_Minimo, i.Ubicacion, i.Peso,i.Tipo FROM Proveedores AS p RIGHT JOIN Insumos AS i ON p.ID = i.Proveedor WHERE i.Codigo <> '-1';";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            System.Data.DataTable banco = new System.Data.DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView1.DataSource = banco;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[8].HeaderText = "Unidad de Medida";
            dataGridView1.Columns[9].HeaderText = "Costo/Unitario";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[10].HeaderText = "Cantidad en Stock";
            dataGridView1.Columns[11].HeaderText = "Stock Minimo";

        }

        public void cargarInsumosStock()
        {
            string query = "SELECT * FROM Insumos WHERE Stock_Minimo >= Cantidad_Stock AND Codigo <> '-1'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView2.DataSource = banco;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[8].HeaderText = "Unidad de Medida";
            dataGridView2.Columns[9].HeaderText = "Costo/Unitario";
            dataGridView2.Columns[9].DefaultCellStyle.Format = "c";
            dataGridView2.Columns[10].HeaderText = "Cantidad en Stock";
            dataGridView2.Columns[11].HeaderText = "Stock Minimo";

        }

        public void buscarInsumoStock()
        {
            string query = "SELECT * FROM Insumos WHERE Stock_Minimo >= Cantidad_Stock ";
            if (!txtCodigo.Text.Equals(""))
            {                
                query += " AND ";
                query += "Codigo LIKE '%" + txtCodigo.Text + "%'";
            }
            if (!txtClase.Text.Equals(""))
            {
                query += " AND ";
                query += "Clase LIKE '%" + txtClase.Text + "%'";
            }
            if (!txtMarca.Text.Equals(""))
            {
                query += " AND ";
                query += "Marca LIKE '%" + txtMarca.Text + "%'";
            }
            if (!txtModelo.Text.Equals(""))
            {
                query += " AND ";               
                query += "Modelo LIKE '%" + txtModelo.Text + "%'";
            }
            if (!txtSerial.Text.Equals(""))
            {
                query += " AND ";
                query += "Serial LIKE '%" + txtSerial.Text + "%'";
            }
            if (!txtDescripcion.Text.Equals(""))
            {
                query += " AND ";
                query += "Descripcion LIKE '%" + txtDescripcion.Text + "%'";
            }
            if (!txtProveedor.Text.Equals(""))
            {
                query += " AND ";
                query += "Proveedor LIKE '%" + txtProveedor.Text + "%'";
            }
            if (!txtUnidad_de_Medida.Text.Equals(""))
            {
                query += " AND ";
                query += "Unidad_Medida LIKE '%" + txtUnidad_de_Medida.Text + "%'";
            }
            if (!txtCosto_Unidad.Text.Equals(""))
            {
                query += " AND ";
                query += "Costo_Unitario LIKE '%" + txtCosto_Unidad.Text + "%'";
            }
            if (!txtCantidad_Stock.Text.Equals(""))
            {
                query += " AND ";
                query += "Cantidad_Stock LIKE '%" + txtCantidad_Stock.Text + "%'";
            }
            if (!txtStock_Minimo.Text.Equals(""))
            {
                query += " AND ";
                query += "Stock_Minimo LIKE '%" + txtStock_Minimo.Text + "%'";
            }
            if (!txtPeso.Text.Equals(""))
            {
                query += " AND ";
                query += "Peso LIKE '%" + txtPeso.Text + "%'";
            }
            if (!txtUbicacion.Text.Equals(""))
            {
                query += " AND ";
                query += "Ubicacion LIKE '%" + txtUbicacion.Text + "%'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView2.DataSource = supervisores;
            dataGridView2.Columns[0].Visible = false;
        }

        public void cargarMaquinaria()
        {
            string query = "SELECT ID,(Tipo + '/' + Modelo + '/' + Placa) As Maquinaria FROM Maquinarias";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.DisplayMember = "Maquinaria";
            comboBox1.ValueMember = "ID";
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarProveedores()
        {
            string query = "SELECT * FROM Proveedores";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtProveedor.DataSource = ds.Tables[0];
            txtProveedor.DisplayMember = "Proveedor";
            txtProveedor.ValueMember = "ID";
            txtProveedor.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtProveedor.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox2.DataSource = ds.Tables[0];
            comboBox2.DisplayMember = "Proveedor";
            comboBox2.ValueMember = "ID";
            comboBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void buscarInsumo()
        {
            string query = "SELECT * FROM Insumos ";
            int i = 0;
            if (!txtCodigo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Codigo LIKE '%" + txtCodigo.Text + "%'";
            }
            if (!txtClase.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Clase LIKE '%" + txtClase.Text + "%'";
            }
            if (!txtMarca.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Marca LIKE '%" + txtMarca.Text + "%'";
            }
            if (!txtModelo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Modelo LIKE '%" + txtModelo.Text + "%'";
            }
            if (!txtSerial.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Serial LIKE '%" + txtSerial.Text + "%'";
            }
            if (!txtDescripcion.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Descripcion LIKE '%" + txtDescripcion.Text + "%'";
            }
            if (!txtProveedor.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Proveedor LIKE '%" + txtProveedor.Text + "%'";
            }
            if (!txtUnidad_de_Medida.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Unidad_Medida LIKE '%" + txtUnidad_de_Medida.Text + "%'";
            }
            if (!txtCosto_Unidad.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Costo_Unitario LIKE '%" + txtCosto_Unidad.Text + "%'";
            }
            if (!txtCantidad_Stock.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Cantidad_Stock LIKE '%" + txtCantidad_Stock.Text + "%'";
            }
            if (!txtStock_Minimo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Stock_Minimo LIKE '%" + txtStock_Minimo.Text + "%'";
            }
            if (!txtPeso.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Peso LIKE '%" + txtPeso.Text + "%'";
            }
            if (!txtUbicacion.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Ubicacion LIKE '%" + txtUbicacion.Text + "%'";
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

        public frmInsumos(int tipo)
        {            
            InitializeComponent();
            tipousuario = tipo;
            cargarInsumos();
            cargarMaquinaria();
            cargarProveedores();
            comboBox2.Enabled = false;
            txtProveedor.SelectedItem = null;
            comboBox1.SelectedItem = null;
            comboBox2.SelectedItem = null;
            if (tipousuario == 1)
            {
                dataGridView1.Columns[9].Visible = false;
                textBox3.Visible = false;
                txtCosto.Visible = false;
                txtCosto_Unidad.Visible = false;
                label8.Visible = false;
                label19.Visible = false;
                label17.Visible = false;
            }
        }

        public void agregarInsumo()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Insumos(Codigo,Clase,Marca,Modelo,Serial,Descripcion,Proveedor,Unidad_Medida,Costo_Unitario,Cantidad_Stock,Stock_Minimo,Ubicacion,Peso,Tipo) VALUES (@Codigo,@Clase,@Marca,@Modelo,@Serial,@Descripcion,@Proveedor,@Unidad_Medida,@Costo_Unitario,@Cantidad_Stock,@Stock_Minimo,@Ubicacion,@Peso,@Tipo)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
                cmd.Parameters.Add("@Clase", OleDbType.VarChar).Value = txtClase.Text;
                cmd.Parameters.Add("@Marca", OleDbType.VarChar).Value = txtMarca.Text;
                cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = txtModelo.Text;
                cmd.Parameters.Add("@Serial", OleDbType.VarChar).Value = txtSerial.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = txtProveedor.SelectedValue;
                cmd.Parameters.Add("@Unidad_Medida", OleDbType.VarChar).Value = txtUnidad_de_Medida.Text;
                cmd.Parameters.Add("@Costo_Unitario", OleDbType.VarChar).Value = txtCosto_Unidad.Text;
                cmd.Parameters.Add("@Cantidad_Stock", OleDbType.VarChar).Value = txtCantidad_Stock.Text;
                cmd.Parameters.Add("@Stock_Minimo", OleDbType.VarChar).Value = txtStock_Minimo.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = txtUbicacion.Text;
                cmd.Parameters.Add("@Peso", OleDbType.VarChar).Value = txtPeso.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Insumo agregado.");
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
            string query = "SELECT MAX(id) FROM Insumos";
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

        public void agregarInsumoMaquina(int id)
        {
            for (int i = 0; i < maquinarias.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO insumoMaquinas(Insumo,Maquina) VALUES (@Insumo,@Maquina)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = id;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = maquinarias[i];
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
            }
        }

        public void modificarInsumo()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Insumos SET Codigo=@Codigo,Clase=@Clase,Marca=@Marca,Modelo=@Modelo,Serial=@Serial,Descripcion=@Descripcion,Proveedor=@Proveedor,Unidad_Medida=@Unidad_Medida,Costo_Unitario=@Costo_Unitario,Cantidad_Stock=@Cantidad_Stock,Stock_Minimo=@Stock_Minimo,Ubicacion=@Ubicacion,Peso=@Peso,Tipo=@Tipo WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Codigo", OleDbType.VarChar).Value = txtCodigo.Text;
                cmd.Parameters.Add("@Clase", OleDbType.VarChar).Value = txtClase.Text;
                cmd.Parameters.Add("@Marca", OleDbType.VarChar).Value = txtMarca.Text;
                cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = txtModelo.Text;
                cmd.Parameters.Add("@Serial", OleDbType.VarChar).Value = txtSerial.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = txtProveedor.SelectedValue;
                cmd.Parameters.Add("@Unidad_Medida", OleDbType.VarChar).Value = txtUnidad_de_Medida.Text;
                cmd.Parameters.Add("@Costo_Unitario", OleDbType.VarChar).Value = txtCosto_Unidad.Text;
                cmd.Parameters.Add("@Cantidad_Stock", OleDbType.VarChar).Value = txtCantidad_Stock.Text;
                cmd.Parameters.Add("@Stock_Minimo", OleDbType.VarChar).Value = txtStock_Minimo.Text;
                cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = txtUbicacion.Text;
                cmd.Parameters.Add("@Peso", OleDbType.VarChar).Value = txtPeso.Text;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Insumo modificado.");
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
            if (txtTipo.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar el tipo de insumo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if(txtCodigo.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar el codigo del insumo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (txtProveedor.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar proveedor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                agregarInsumo();
                int id = getMaxID();
                agregarInsumoMaquina(id);
                reiniciarTablero();
                cargarInsumos();
                cargarInsumosStock();
                if (tipousuario == 1)
                {
                    dataGridView1.Columns[9].Visible = false;
                    dataGridView2.Columns[9].Visible = false;
                }
            }            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void getMaquinas(string id) {
            listBox1.Items.Clear();
            maquinarias.Clear();
            string query = "SELECT i.ID, i.Insumo, (m.Tipo + ' /' + m.Marca + '/' + m.Placa) As Maquinaria, m.ID FROM Maquinarias AS m INNER JOIN insumoMaquinas AS i ON m.ID = i.Maquina WHERE i.Insumo = " + id;
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
                    listBox1.Items.Add(myReader.GetString(2));
                    maquinarias.Add(myReader.GetInt32(3).ToString());
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {            
            txtCodigo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtClase.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtMarca.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtModelo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtSerial.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtDescripcion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtProveedor.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtUnidad_de_Medida.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtCosto_Unidad.Text = String.Format("{0:c}",dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString());
            txtCosto.Text = String.Format("{0:c}",dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString());
            comboBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox3.Text = String.Format("{0:c}",dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value);
            txtCantidad_Stock.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            txtStock_Minimo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            txtUbicacion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            txtPeso.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            txtTipo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();            
            getMaquinas(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());            
        }

        public void reiniciarTablero() {
            txtCodigo.Text = "";
            txtClase.Text = "";
            txtMarca.Text = "";
            txtModelo.Text = "";
            txtSerial.Text = "";
            txtDescripcion.Text = "";
            txtProveedor.Text = "";
            txtUnidad_de_Medida.Text = "";
            txtCosto_Unidad.Text = "";
            txtCantidad_Stock.Text = "";
            txtStock_Minimo.Text = "";
            txtUbicacion.Text = "";
            txtCosto.Text = "";
            textBox3.Text = "";
            textBox2.Text = "";
            textBox1.Text = "1";
            txtPeso.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            listBox1.Items.Clear();
            maquinarias.Clear();

        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!comboBox1.Text.Equals(""))
            {
                listBox1.Items.Add(comboBox1.Text);
                maquinarias.Add(comboBox1.SelectedValue.ToString());                                    
            }
        }        

        private void button2_Click(object sender, EventArgs e)
        {

            maquinarias.RemoveAt(listBox1.SelectedIndex);
            listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el insumo " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                eliminarInsumoMaquinarias();    
                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Insumos WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Insumo eliminado.");
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
                cargarInsumos();
                cargarInsumosStock();   
                reiniciarTablero();
                if (tipousuario == 1)
                {
                    dataGridView1.Columns[9].Visible = false;
                    dataGridView2.Columns[9].Visible = false;
                }
            }
        }

        public void eliminarInsumoMaquinarias() {
            for (int i = 0; i < maquinarias.Count; i++)
            {
             string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM insumoMaquinas WHERE Insumo = " + id);
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
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            modificarInsumo();
            eliminarInsumoMaquinarias();            
            agregarInsumoMaquina(Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
            cargarInsumos();
            cargarInsumosStock();
            reiniciarTablero();
            if (tipousuario == 1)
            {
                dataGridView1.Columns[9].Visible = false;
                dataGridView2.Columns[9].Visible = false;
            }
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                buscarInsumo();
                if (tipousuario == 1)
                {
                    dataGridView1.Columns[9].Visible = false;
                }
            }
            else
            {
                buscarInsumoStock();
                if (tipousuario == 1)
                {
                    dataGridView2.Columns[9].Visible = false;
                }
            }

        }

        public void busquedaMaquina()
        {
            string query = "SELECT i.* FROM Insumos AS i INNER JOIN insumoMaquinas AS m ON i.ID = m.Insumo WHERE m.Maquina = " + comboBox1.SelectedValue;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView1.DataSource = banco;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[8].HeaderText = "Unidad de Medida";
            dataGridView1.Columns[9].HeaderText = "Costo/Unitario";
            dataGridView1.Columns[10].HeaderText = "Cantidad en Stock";
            dataGridView1.Columns[11].HeaderText = "Stock Minimo";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            busquedaMaquina();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                cargarInsumos();
                reiniciarTablero();
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                cargarInsumosStock();
                if (tipousuario == 1)
                {
                    dataGridView2.Columns[9].Visible = false;
                }
                reiniciarTablero();
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtCodigo.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtClase.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtMarca.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtModelo.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[4].Value.ToString();
            txtSerial.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtDescripcion.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtProveedor.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtUnidad_de_Medida.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtCosto_Unidad.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[9].Value.ToString();
            txtCantidad_Stock.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[10].Value.ToString();
            txtStock_Minimo.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[11].Value.ToString();
            txtUbicacion.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[12].Value.ToString();
            txtPeso.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[13].Value.ToString();
            getMaquinas(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());    
        }

        //public void imprimirInsumos() {
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;
        //    object misValue = System.Reflection.Missing.Value;

        //    xlApp = new Excel.Application();
        //    xlWorkBook = xlApp.Workbooks.Add(misValue);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        //    int i = 0;
        //    int j = 0;

        //    for (i = 0; i <= dataGridView1.RowCount - 1; i++)
        //    {
        //        for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
        //        {
        //            DataGridViewCell cell = dataGridView1[j, i];
        //            xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
        //        }
        //    }

        //    xlWorkBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //    xlWorkBook.Close(true, misValue, misValue);
        //    xlApp.Quit();
            
        //    releaseObject(xlWorkSheet);
        //    releaseObject(xlWorkBook);
        //    releaseObject(xlApp);

        //    MessageBox.Show("Excel file created , you can find the file c:\\csharp.net-informations.xls");
        //}
        //private void releaseObject(object obj)
        //{
        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //        MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
        //    }
        //    finally
        //    {
        //        GC.Collect();
        //    }
        //}

        public void imprimirInsumos2(DataGridView data) {             
             if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                XcelApp.Application.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 2; i < data.Columns.Count + 1; i++)
                {
                    XcelApp.Cells[2, i] = data.Columns[i - 1].HeaderText;
                }
 
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 1; j < data.Columns.Count; j++)
                    {
                        XcelApp.Cells[i + 3, j + 1] = data.Rows[i].Cells[j].Value.ToString();
                        if (i==0)
                        {
                            excelCellrange = XcelApp.Range[XcelApp.Cells[i+2, 2], XcelApp.Cells[i+2, data.Columns.Count]];
                            excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                excelCellrange = XcelApp.Range[XcelApp.Cells[2, 2], XcelApp.Cells[data.Rows.Count+2, data.Columns.Count]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
                XcelApp.Columns.AutoFit();
                XcelApp.Visible = true;
            }        
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)            
                imprimirInsumos2(dataGridView1);            
            else
                imprimirInsumos2(dataGridView2);
        }

        private void button5_Click(object sender, EventArgs e)
        {
        }

        public void comprarInsumo()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Insumos SET Costo_Unitario=@Costo_Unitario,Cantidad_Stock=@Cantidad_Stock WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cost_Unitario", OleDbType.VarChar).Value = Int32.Parse(txtCosto.Text);
                cmd.Parameters.Add("@Cantidad_Stock", OleDbType.VarChar).Value = Int32.Parse(textBox2.Text) + Int32.Parse(textBox1.Text);
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Insumo comprado.");
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

        public void proveedorInsumo()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO proveedorInsumo(Proveedor,Insumo,Cantidad,Fecha,Costo) VALUES (@Proveedor,@Insumo,@Cantidad,@Fecha,@Costo)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year;
                cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = textBox3.Text;
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


        private void button7_Click(object sender, EventArgs e)
        {
            comprarInsumo();
            proveedorInsumo();
            cargarInsumos();
            reiniciarTablero();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string costo = txtCosto.Text.Replace("$", "").Replace(".", "");
            if(!textBox1.Text.Equals("") && !costo.Equals(""))
                textBox3.Text = String.Format("{0:c}",Double.Parse(costo) * Double.Parse(textBox1.Text));
        }

        private void txtCosto_TextChanged(object sender, EventArgs e)
        {
            string costo = txtCosto.Text.Replace("$", "").Replace(".","");
            if (!costo.Equals(""))
                textBox3.Text = String.Format("{0:c}", Double.Parse(costo) * Double.Parse(textBox1.Text));
        }
    }
}
