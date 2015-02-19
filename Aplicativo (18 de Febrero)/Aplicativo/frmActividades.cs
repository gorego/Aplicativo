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
    public partial class frmActividades : Form
    {
        //Base de datos.
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        double suma = 0;

        public void cargarFormatos()
        {
            string query = "SELECT * FROM Formatos";
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
                    checkedListBox1.Items.Add(myReader.GetString(1));                    
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

        public void cargarActividades()
        {
            string query = "SELECT a.ID,a.Actividad, a.Unidad_de_Medida,SUM(i.Costo_Unitario*ia.Cantidad), d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza FROM (Actividades AS a INNER JOIN (Insumos AS i INNER JOIN insumoActividades AS ia ON i.ID = ia.Actividad) ON a.ID = ia.Insumo) INNER JOIN Departamentos AS d ON d.ID = a.Departamento GROUP BY a.ID,a.Actividad, a.Unidad_de_Medida,d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].HeaderText = "Unidad de Medida";
            dataGridView1.Columns[3].DefaultCellStyle.Format = "c";
            dataGridView1.Columns[3].HeaderText = "Costo/Unidad";
            dataGridView1.Columns[5].HeaderText = "Tipo";
            dataGridView1.Columns[6].HeaderText = "Condiciones Minimas";
            dataGridView1.Columns[7].HeaderText = "Descpricion";
        }

        public void buscarInsumos()
        {
            string query = "SELECT a.ID,a.Actividad, a.Unidad_de_Medida,SUM(i.Costo_Unitario*ia.Cantidad), d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza FROM (Actividades AS a INNER JOIN (Insumos AS i INNER JOIN insumoActividades AS ia ON i.ID = ia.Actividad) ON a.ID = ia.Insumo) INNER JOIN Departamentos AS d ON d.ID = a.Departamento WHERE ia.Actividad = " + comboBox1.SelectedValue + " GROUP BY a.ID,a.Actividad, a.Unidad_de_Medida,d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza";          
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            dataGridView1.DataSource = supervisores;
            dataGridView1.Columns[0].Visible = false;
        }

        public void buscarActividades()
        {
            string query = "SELECT a.ID,a.Actividad, a.Unidad_de_Medida,SUM(i.Costo_Unitario*ia.Cantidad), d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza FROM (Actividades AS a INNER JOIN (Insumos AS i INNER JOIN insumoActividades AS ia ON i.ID = ia.Actividad) ON a.ID = ia.Insumo) INNER JOIN Departamentos AS d ON d.ID = a.Departamento ";
            int i = 0;
            if(!txtActividad.Text.Equals("")){
                if(i!=0)
                    query+=" AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Actividad LIKE '%"+txtActividad.Text+"%'";
            }
            if (!txtUnidad.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Unidad_de_Medida LIKE '%" + txtUnidad.Text + "%'";
            }
            //if (!txtCosto.Text.Equals(""))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "Costo_Unidad LIKE '%" + txtCosto.Text + "%'";
            //}
            if (!txtDepartamento.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "d.Departamento LIKE '%" + txtDepartamento.Text + "%'";
            }
            if (!txtDescripcion.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Descripcion_Actividad LIKE '%" + txtDescripcion.Text + "%'";
            }
            if (!txtTipo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Tipo_Actividad LIKE '%" + txtTipo.Text + "%'";
            }
            if (!txtRodilla.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Rodilla LIKE '%" + txtRodilla.Text + "%'";
            }
            if (!txtCintura.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Cintura LIKE '%" + txtCintura.Text + "%'";
            }
            if (!txtCabeza.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "a.Cabeza LIKE '%" + txtCabeza.Text + "%'";
            }
            query += "GROUP BY a.ID,a.Actividad, a.Unidad_de_Medida,d.Departamento,a.Tipo_Actividad,a.Condicion_Minima,a.Descripcion_Actividad,a.Rodilla,a.Cintura,a.Cabeza";
            //if (checkedListBox1.GetItemCheckState(0).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF001 LIKE 'Si'";
            //}
            //if (checkedListBox1.GetItemCheckState(1).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF002 LIKE 'Si'";
            //}
            //if (checkedListBox1.GetItemCheckState(2).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF003 LIKE 'Si'";
            //}
            //if (checkedListBox1.GetItemCheckState(3).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF004 LIKE 'Si'";
            //}
            //if (checkedListBox1.GetItemCheckState(4).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF005 LIKE 'Si'";
            //}
            //if (checkedListBox1.GetItemCheckState(5).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF006 LIKE 'Si'";
            //}
            //if (checkedListBox1.GetItemCheckState(6).ToString().Equals("Checked"))
            //{
            //    if (i != 0)
            //        query += " AND ";
            //    else
            //        query += "WHERE ";
            //    i++;
            //    query += "ADF007 LIKE 'Si'";
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

        public void cargarDepartamentos()
        {
            string query = "SELECT * FROM Departamentos";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable departamentos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtDepartamento.DataSource = ds.Tables[0];
            txtDepartamento.DisplayMember = "Departamento";
            txtDepartamento.ValueMember = "ID";            
        }

        public void cargarInsumos()
        {
            string query = "SELECT ID, Modelo As Insumo FROM Insumos";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable departamentos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView2.Columns[2].DefaultCellStyle.Format = "c";
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.DisplayMember = "Insumo";
            comboBox1.ValueMember = "ID";
        }

        public void agregarActividad()
        {
            conn.ConnectionString = connectionString;
            //OleDbCommand cmd = new OleDbCommand("INSERT INTO Actividades (Actividad,Unidad_de_Medida,Costo_Unidad,Departamento,Tipo_Actividad, Descripcion_Actividad,Condicion_Minima,Rodilla,Cintura,Cabeza,ADF001,ADF002,ADF003,ADF004,ADF005,ADF006,ADF007) VALUES (@Actividad,@Unidad,@Costo,@Dept,@Tipo,@Descripcion,@Condicion,@Rodilla,@Cintura,@Cabeza,@ADF001,@ADF002,@ADF003,@ADF004,@ADF005,@ADF006,@ADF007)");
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Actividades (Actividad,Unidad_de_Medida,Departamento,Tipo_Actividad, Descripcion_Actividad,Condicion_Minima,Rodilla,Cintura,Cabeza) VALUES (@Actividad,@Unidad,@Dept,@Tipo,@Descripcion,@Condicion,@Rodilla,@Cintura,@Cabeza)");
            cmd.Connection = conn;
            conn.Open();            
            if (conn.State == ConnectionState.Open)
            {                
                cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = txtActividad.Text;
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidad.Text;
                //cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtCosto.Text;
                cmd.Parameters.Add("@Dept", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Condicion", OleDbType.VarChar).Value = txtCondiciones.Text;
                cmd.Parameters.Add("@Rodilla", OleDbType.VarChar).Value = txtRodilla.Text;
                cmd.Parameters.Add("@Cintura", OleDbType.VarChar).Value = txtCintura.Text;
                cmd.Parameters.Add("@Cabeza", OleDbType.VarChar).Value = txtCabeza.Text;
                //if (checkedListBox1.GetItemCheckState(0).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF001", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF001", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(1).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF002", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF002", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(2).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF003", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF003", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(3).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF004", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF004", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(4).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF005", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF005", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(5).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF006", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF006", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(6).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF007", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF007", OleDbType.VarChar).Value = "No";
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Actividad agregada.");
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

        public void modificarActividad()
        {
            conn.ConnectionString = connectionString;
            //OleDbCommand cmd = new OleDbCommand("UPDATE Actividades SET Actividad=@Actividad,Unidad_de_Medida=@Unidad,Costo_Unidad=@Costo,Departamento=@Dept,Tipo_Actividad=@Tipo,Descripcion_Actividad=@Descripcion,Condicion_Minima=@Condicion_Minima,Rodilla=@Rodilla,Cintura=@Cintura,Cabeza=@Cabeza,ADF001=@ADF001,ADF002=@ADF002,ADF003=@ADF003,ADF004=@ADF004,ADF005=@ADF005,ADF006=@ADF006,ADF007=@ADF007 WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            OleDbCommand cmd = new OleDbCommand("UPDATE Actividades SET Actividad=@Actividad,Unidad_de_Medida=@Unidad,Departamento=@Dept,Tipo_Actividad=@Tipo,Descripcion_Actividad=@Descripcion,Condicion_Minima=@Condicion_Minima,Rodilla=@Rodilla,Cintura=@Cintura,Cabeza=@Cabeza WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = txtActividad.Text;
                cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = txtUnidad.Text;
                //cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtCosto.Text;
                cmd.Parameters.Add("@Dept", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
                cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = txtTipo.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Condicion_Minima", OleDbType.VarChar).Value = txtCondiciones.Text;
                cmd.Parameters.Add("@Rodilla", OleDbType.VarChar).Value = txtRodilla.Text;
                cmd.Parameters.Add("@Cintura", OleDbType.VarChar).Value = txtCintura.Text;
                cmd.Parameters.Add("@Cabeza", OleDbType.VarChar).Value = txtCabeza.Text;
                //if (checkedListBox1.GetItemCheckState(0).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF001", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF001", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(1).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF002", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF002", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(2).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF003", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF003", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(3).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF004", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF004", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(4).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF005", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF005", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(5).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF006", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF006", OleDbType.VarChar).Value = "No";
                //if (checkedListBox1.GetItemCheckState(6).ToString().Equals("Checked"))
                //    cmd.Parameters.Add("@ADF007", OleDbType.VarChar).Value = "Si";
                //else
                //    cmd.Parameters.Add("@ADF007", OleDbType.VarChar).Value = "No";
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Actividad modficiada.");
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

        public void eliminarActividad() {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la actividad " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {

                    string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("DELETE FROM Actividades WHERE id = " + id);
                    cmd.Connection = conn;
                    conn.Open();

                    if (conn.State == ConnectionState.Open)
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Actividad eliminada.");
                            conn.Close();
                            cargarActividades();
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
                MessageBox.Show("Favor seleccionar el nombre de la actividad.", "Error");
            }
        }

        public frmActividades()
        {
            InitializeComponent();
            cargarActividades();
            cargarDepartamentos();
            cargarInsumos();
            cargarFormatos();
            comboBox1.SelectedIndex = -1;
            txtDepartamento.SelectedIndex = -1;
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            Variables.cargar2(comboBox2, "SELECT Clase From Insumos GROUP BY Clase", "Clase");
            Variables.cargar2(comboBox3, "SELECT Marca From Insumos GROUP BY Marca", "Marca");
        }

        public void reiniciarTabler() {
            txtActividad.Text = "";
            txtUnidad.Text = "";
            txtCosto.Text = "";
            txtDepartamento.Text = "";
            txtTipo.Text = "";
            txtCondiciones.Text = "";
            txtDescripcion.Text = "";
            txtRodilla.Text = "";
            txtCintura.Text = "";
            txtCabeza.Text = "";
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {               
                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            } 
            while (dataGridView2.Rows.Count != 0)
            {
                dataGridView2.Rows.RemoveAt(0);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public int getMaxID()
        {
            string query = "SELECT MAX(id) FROM Actividades";
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

        public void agregarInsumoActividad(int id)
        {
            //if (dataGridView2.Rows.Count == 0)
            //{
            //    conn.ConnectionString = connectionString;
            //    OleDbCommand cmd = new OleDbCommand("INSERT INTO insumoActividades(Actividad,Insumo,Cantidad) VALUES (@Actividad,@Insumo,@Cantidad)");
            //    cmd.Connection = conn;
            //    conn.Open();
            //    if (conn.State == ConnectionState.Open)
            //    {
            //        cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = 1;
            //        cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = id;
            //        cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = 0;
            //        try
            //        {
            //            cmd.ExecuteNonQuery();
            //            conn.Close();
            //        }
            //        catch (OleDbException ex)
            //        {
            //            MessageBox.Show(ex.Source);
            //            conn.Close();
            //        }
            //    }
            //}
            if (dataGridView2.Rows.Count != 0)
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("INSERT INTO insumoActividades(Actividad,Insumo,Cantidad) VALUES (@Actividad,@Insumo,@Cantidad)");
                    cmd.Connection = conn;
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = dataGridView2.Rows[i].Cells[0].Value;
                        cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = id;
                        cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = dataGridView2.Rows[i].Cells[3].Value;
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
        }

        public void agregarFormatosActividad(int id)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemCheckState(i).ToString().Equals("Checked")){
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("INSERT INTO formatosActividad(Actividad,Formato) VALUES (@Actividad,@Formato)");
                    cmd.Connection = conn;
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        cmd.Parameters.Add("@Actividad", OleDbType.VarChar).Value = id;
                        cmd.Parameters.Add("@Formato", OleDbType.VarChar).Value = checkedListBox1.Items[i].ToString();
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
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!txtActividad.Text.Equals(""))
            {
                if (!txtUnidad.Text.Equals(""))
                {
                    if (!txtTipo.Text.Equals(""))
                    {
                        if (!txtDepartamento.Text.Equals(""))
                        {
                            if (!txtRodilla.Text.Equals("") || !txtCintura.Text.Equals("") || !txtCabeza.Text.Equals(""))
                            {
                                agregarActividad();
                                int act = getMaxID();
                                agregarInsumoActividad(act);
                                agregarFormatosActividad(act);
                                cargarActividades();
                                reiniciarTabler();
                            }
                            else
                                MessageBox.Show("Favor ingresar Rodilla, Cintura, Cabeza.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else
                            MessageBox.Show("Favor ingreresar el tipo de actividad.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                        MessageBox.Show("Favor ingreresar el tipo de actividad.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                    MessageBox.Show("Favor ingreresar la unidad de medida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else
                MessageBox.Show("Favor ingreresar nombre de la actividad.","Error",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            eliminarActividad();
            reiniciarTabler();
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {
            buscarActividades();
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTabler();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmOrdenes newFrm = new frmOrdenes("Actividad", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            getInsumos(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            getFormatos(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            txtActividad.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();            
            txtUnidad.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            //txtCosto.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            txtCosto.Text = string.Format("{0:c}",suma.ToString());
            txtDepartamento.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();    
            txtTipo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtCondiciones.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtDescripcion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtRodilla.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtCintura.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            txtCabeza.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();           
            //for (int i = 11; i < 18; i++)
            //{
            //    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[i].Value.ToString().Equals("Si"))
            //        checkedListBox1.SetItemCheckState(i - 11, CheckState.Checked);
            //    else
            //        checkedListBox1.SetItemCheckState(i - 11, CheckState.Unchecked);
            //}
        }

        public void eliminarInsumoActividades()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM insumoActividades WHERE Insumo = " + id);
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

        public void eliminarFormatoActividades()
        {
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM formatosActividad WHERE Actividad = " + id);
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (!txtActividad.Text.Equals(""))
            {
                if (!txtUnidad.Text.Equals(""))
                {
                    if (!txtTipo.Text.Equals(""))
                    {
                        if (!txtDepartamento.Text.Equals(""))
                        {
                            if (!txtRodilla.Text.Equals("") || !txtCintura.Text.Equals("") || !txtCabeza.Text.Equals(""))
                            {
                                modificarActividad();
                                eliminarInsumoActividades();
                                agregarInsumoActividad(Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
                                eliminarFormatoActividades();
                                agregarFormatosActividad(Int32.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()));
                                cargarActividades();
                                reiniciarTabler();
                            }
                            else
                                MessageBox.Show("Favor ingresar Rodilla, Cintura, Cabeza.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else
                            MessageBox.Show("Favor ingreresar el tipo de actividad.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                        MessageBox.Show("Favor ingreresar el tipo de actividad.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                    MessageBox.Show("Favor ingreresar la unidad de medida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            else
                MessageBox.Show("Favor ingreresar nombre de la actividad.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        public void getInsumos(string id)
        {
            suma = 0;
            reiniciarTabler();
            string query = "SELECT i.ID, i.Insumo, a.ID,(a.Codigo + ' ' + a.Marca + ' ' + a.Modelo) As Insumo, a.Costo_Unitario, i.Cantidad FROM Insumos AS a INNER JOIN insumoActividades AS i ON a.ID = i.Actividad WHERE i.Insumo = " + id;
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
                    double cost = double.Parse(myReader.GetDouble(4).ToString()) * myReader.GetDouble(5);
                    dataGridView2.Rows.Add();
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[0].Value = myReader.GetInt32(2);
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1].Value = myReader.GetString(3);
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[2].Value = cost;
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[3].Value = myReader.GetDouble(5);
                    suma += cost;
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

        public void getFormatos(string id)
        {
            string query = "SELECT i.ID, i.Formato FROM Actividades AS a INNER JOIN formatosActividad AS i ON a.ID = i.Actividad WHERE i.Actividad = " + id;
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
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if ((string)checkedListBox1.Items[i] == myReader.GetString(1))
                        {
                            checkedListBox1.SetItemChecked(i, true);
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

        public void getInsumo(string id) {            
            string query = "SELECT ID,(Codigo + ' ' + Marca + ' ' + Modelo) As Insumo, Costo_Unitario FROM Insumos WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            try
            {
                conn.ConnectionString = connectionString;
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataReader reader = cmd.ExecuteReader();                
                if (reader.Read())
                {
                    dataGridView2.Rows.Add();                    
                    dataGridView2.Rows[dataGridView2.Rows.Count-1].Cells[0].Value = reader.GetInt32(0);
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1].Value = reader.GetString(1);
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[2].Value = ((reader.GetDouble(2)) * double.Parse(textBox1.Text));
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[3].Value = textBox1.Text;
                    suma += (reader.GetDouble(2) * double.Parse(textBox1.Text));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }        

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals("")) {
                MessageBox.Show("Favor ingresar cantidad de insumo.", "Error");
            }
            else
            {
                getInsumo(comboBox1.SelectedValue.ToString());
                txtCosto.Text = suma.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            suma -= double.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString());
            txtCosto.Text = suma.ToString();
            dataGridView2.Rows.Remove(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex]);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            buscarInsumos();
        }

        private void txtRodilla_MouseLeave(object sender, EventArgs e)
        {
        }

        private void txtRodilla_TextChanged(object sender, EventArgs e)
        {
            //int n;
            //bool isNumeric = int.TryParse(txtRodilla.Text, out n);
            //if (!isNumeric)
            //{
            //    txtRodilla.Text = "1";
            //    MessageBox.Show("Favor ingreresar un numero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //}
        }

        private void txtCintura_TextChanged(object sender, EventArgs e)
        {
            //int n;
            //bool isNumeric = int.TryParse(txtCintura.Text, out n);
            //if (!isNumeric)
            //{
            //    txtCintura.Text = "1";
            //    MessageBox.Show("Favor ingreresar un numero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //}
        }

        private void txtCabeza_TextChanged(object sender, EventArgs e)
        {
            //int n;
            //bool isNumeric = int.TryParse(txtCabeza.Text, out n);
            //if (!isNumeric)
            //{
            //    txtCabeza.Text = "1";
            //    MessageBox.Show("Favor ingreresar un numero.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //}
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmFormatos newFrm = new frmFormatos();
            newFrm.Show();
        }

        public void imprimirActividades(DataGridView data)
        {
            if (data.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                XcelApp.Application.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                for (int i = 1; i < data.Columns.Count + 1; i++)
                {
                    XcelApp.Cells[2, i + 1] = data.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        XcelApp.Cells[i + 3, j + 2] = data.Rows[i].Cells[j].Value.ToString();
                        if (i == 0)
                        {
                            excelCellrange = XcelApp.Range[XcelApp.Cells[i + 2, 2], XcelApp.Cells[i + 2, data.Columns.Count + 1]];
                            excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            excelCellrange.Font.Color = System.Drawing.Color.White;
                        }
                    }
                }
                excelCellrange = XcelApp.Range[XcelApp.Cells[2, 2], XcelApp.Cells[data.Rows.Count + 2, data.Columns.Count + 1]];
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
            imprimirActividades(dataGridView1);
        }

        private void checkedListBox1_MouseHover(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Variables.cargar(comboBox1, "SELECT * FROM Insumos WHERE Clase = '" + comboBox2.Text + "'", "Modelo");
            comboBox3.Items.Clear();
            Variables.cargar2(comboBox3, "SELECT Marca FROM Insumos WHERE Clase = '" + comboBox2.Text + "' GROUP BY Marca", "Marca");
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.Equals(""))
                Variables.cargar(comboBox1, "SELECT * FROM Insumos WHERE Marca = '" + comboBox3.Text + "'", "Modelo");
            else
                Variables.cargar(comboBox1, "SELECT * FROM Insumos WHERE Marca = '" + comboBox3.Text + "' AND Clase = '" + comboBox2.Text + "'", "Modelo");
            //Variables.cargar(comboBox1, "SELECT * FROM Insumos WHERE Marca = '" + comboBox3.Text + "'", "Modelo");
            //Variables.cargar2(comboBox2, "SELECT Clase FROM Insumos WHERE Marca = '" + comboBox3.Text + "' GROUP BY Clase", "Clase");
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text.Equals("Dia Hombre"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Jornal");
            }
            else if (comboBox4.Text.Equals("Medida de Area"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Ha");
                txtUnidad.Items.Add("m2");
            }
            else if (comboBox4.Text.Equals("Medida de Cantidad"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Arbol");
                txtUnidad.Items.Add("Plantulas");
            }
            else if (comboBox4.Text.Equals("Medida de Longitud"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Mt");
            }
            else if (comboBox4.Text.Equals("Medida de Peso"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Kg");
            }
            else if (comboBox4.Text.Equals("Medida de Tiempo"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Dias");
                txtUnidad.Items.Add("Horas");
            }
            else if (comboBox4.Text.Equals("Medida de Volumen"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Sacos");
                txtUnidad.Items.Add("Litros");
                txtUnidad.Items.Add("Galones");
                txtUnidad.Items.Add("m3");
            }
            else if (comboBox4.Text.Equals("Recorridos"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Recorridos");
            }
            else if (comboBox4.Text.Equals("Todos"))
            {
                txtUnidad.Items.Clear();
                txtUnidad.Items.Add("Jornal");
                txtUnidad.Items.Add("Ha");
                txtUnidad.Items.Add("m2");
                txtUnidad.Items.Add("Arbol");
                txtUnidad.Items.Add("Plantulas");
                txtUnidad.Items.Add("Mt");
                txtUnidad.Items.Add("Kg");
                txtUnidad.Items.Add("Dias");
                txtUnidad.Items.Add("Horas");
                txtUnidad.Items.Add("Sacos");
                txtUnidad.Items.Add("Litros");
                txtUnidad.Items.Add("Galones");
                txtUnidad.Items.Add("m3");
                txtUnidad.Items.Add("Recorridos");
            }
        }
    }
}
