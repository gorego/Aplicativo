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
using System.Globalization;

namespace Aplicativo
{
    public partial class frmVerMaquinaria : Form
    {
        //Base de datos.
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string id_maquina = "";
        List<string> insumos = new List<string>();
        List<string> cantidad = new List<string>();

        public void cargarMaquinaria()
        {
            string query = "SELECT m.ID, m.Placa, m.Tipo, m.Marca, m.Modelo, e.Estado, m.Ano, m.Ano_Fabricacion,m.Tipo_Combustible, m.Descripcion, p.Propietario, m.aceiteMotor, m.aceiteHidraulico, m.aceiteCajaV, m.aceiteDiferencial FROM propietariosMaquina AS p INNER JOIN (Estados AS e INNER JOIN Maquinarias AS m ON e.ID = m.Estado) ON p.ID = m.Propietario WHERE m.ID = " + id_maquina;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinarias = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(maquinarias);
            dataGridView3.DataSource = maquinarias;
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[6].HeaderText = "Año de Compra";
            dataGridView3.Columns[7].HeaderText = "Año de Fabricación";
            dataGridView3.Columns[8].HeaderText = "Tipo de Combustible";
            dataGridView3.Columns[11].HeaderText = "Ref. Aceite Motor";
            dataGridView3.Columns[12].HeaderText = "Ref. Aceite Hidraulico";
            dataGridView3.Columns[13].HeaderText = "Ref. Aceite de Caja V";
            dataGridView3.Columns[14].HeaderText = "Ref. Aceite de Diferenciales";
            this.Text = dataGridView3.Rows[0].Cells[2].Value.ToString() + " / " + dataGridView3.Rows[0].Cells[3].Value.ToString() + " / " + dataGridView3.Rows[0].Cells[4].Value.ToString();
        }

        //public void cargarMaquinaAceite()
        //{
        //    string query = "SELECT * FROM MaquinaAceite WHERE id_Maquina = '" + id_maquina + "'";
        //    //Ejecutar el query y llenar el GridView.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    DataTable maquinaaceite = new DataTable();
        //    OleDbDataAdapter da2 = new OleDbDataAdapter(cmd);
        //    da2.Fill(maquinaaceite);
        //    dataGridView2.DataSource = maquinaaceite;
        //    for (int i = 0; i < dataGridView2.Rows.Count; i++)
        //    {
        //        dataGridView2.Rows[i].Cells[3].Value = String.Format("{0:c}", double.Parse(dataGridView2.Rows[i].Cells[3].Value.ToString()));
        //    }
        //    dataGridView2.Columns[0].Visible = false;
        //    dataGridView2.Columns[1].Visible = false;
        //}

        //public void cargarMaquinaMantenimiento()
        //{
        //    string query = "SELECT * FROM MaquinaMantenimiento WHERE id_Maquina = '" + id_maquina + "'";
        //    //Ejecutar el query y llenar el GridView.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    DataTable maquinaaceite = new DataTable();
        //    OleDbDataAdapter da2 = new OleDbDataAdapter(cmd);
        //    da2.Fill(maquinaaceite);
        //    dataGridView4.DataSource = maquinaaceite;
        //    for (int i = 0; i < dataGridView4.Rows.Count; i++)
        //    {
        //        dataGridView4.Rows[i].Cells[3].Value = String.Format("{0:c}", double.Parse(dataGridView4.Rows[i].Cells[3].Value.ToString()));
        //    }
        //    dataGridView4.Columns[0].Visible = false;
        //    dataGridView4.Columns[1].Visible = false;
        //}

        public void totalReparacion() { 
            string query = "SELECT SUM(Costo) As Total FROM MaquinaReparacion WHERE id_Maquina = '" + id_maquina + "'";
            //Ejecutar el query y llenar el GridView.
            try
            {
                conn.ConnectionString = connectionString;
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataReader reader = cmd.ExecuteReader();
                if(reader.Read())
                    lblTotalRep.Text = "Total: " + String.Format("{0:c}",reader.GetValue(0));
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

        //public void totalAceite()
        //{
        //    string query = "SELECT SUM(Costo) As Total FROM MaquinaAceite WHERE id_Maquina = '" + id_maquina + "'";
        //    //Ejecutar el query y llenar el GridView.
        //    try
        //    {
        //        conn.ConnectionString = connectionString;
        //        conn.Open();
        //        OleDbCommand cmd = new OleDbCommand(query, conn);
        //        OleDbDataReader reader = cmd.ExecuteReader();
        //        if (reader.Read())
        //            lblCambio.Text = "Total: " + String.Format("{0:c}", reader.GetValue(0));
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {
        //        conn.Close();
        //    }
        //}

        //public void totalMantenimiento()
        //{
        //    string query = "SELECT SUM(Costo) As Total FROM MaquinaMantenimiento WHERE id_Maquina = '" + id_maquina + "'";
        //    //Ejecutar el query y llenar el GridView.
        //    try
        //    {
        //        conn.ConnectionString = connectionString;
        //        conn.Open();
        //        OleDbCommand cmd = new OleDbCommand(query, conn);
        //        OleDbDataReader reader = cmd.ExecuteReader();
        //        if (reader.Read())
        //            label5.Text = "Total: " + String.Format("{0:c}", reader.GetValue(0));
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {
        //        conn.Close();
        //    }
        //}

        public void cargarMaquinaReparacion()
        {
            string query = "SELECT * FROM MaquinaReparacion WHERE id_Maquina = '" + id_maquina + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaaceite = new DataTable();
            OleDbDataAdapter da2 = new OleDbDataAdapter(cmd);
            da2.Fill(maquinaaceite);
            dataGridView1.DataSource = maquinaaceite;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells[3].Value = String.Format("{0:c}", double.Parse(dataGridView1.Rows[i].Cells[3].Value.ToString()));
            }
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Visible = false;
        }

        //public void agregarAceite()
        //{
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand("INSERT INTO MaquinaAceite (id_Maquina,Fecha,Costo,Descripcion) VALUES (@id,@Fecha,@Costo,@Descripcion)");
        //    cmd.Connection = conn;
        //    conn.Open();
        //    if (conn.State == ConnectionState.Open)
        //    {
        //        cmd.Parameters.Add("@id", OleDbType.VarChar).Value = id_maquina;
        //        cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker2.Value.Day.ToString() + "/" + dateTimePicker2.Value.Month.ToString() + "/" + dateTimePicker2.Value.Year.ToString();
        //        cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtCostoAceite.Text;
        //        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcionAceite.Text;
        //        try
        //        {
        //            cmd.ExecuteNonQuery();
        //            MessageBox.Show("Cambio de aceite agregado.");
        //            conn.Close();
        //        }
        //        catch (OleDbException ex)
        //        {
        //            MessageBox.Show(ex.Source);
        //            conn.Close();
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Connection Failed");
        //    }
        //}

        public void agregarDetalles(int reparacion, string detalle, string costo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO reparacionDetalles (Reparacion,Detalle,Costo) VALUES (@Reparacion,@Detalle,@Costo)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Reparacion", OleDbType.VarChar).Value = reparacion;
                cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = detalle;
                cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = costo;
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

        //public void agregarMantenimiento()
        //{
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand("INSERT INTO MaquinaMantenimiento (id_Maquina,Fecha,Costo,Descripcion) VALUES (@id,@Fecha,@Costo,@Descripcion)");
        //    cmd.Connection = conn;
        //    conn.Open();
        //    if (conn.State == ConnectionState.Open)
        //    {
        //        cmd.Parameters.Add("@id", OleDbType.VarChar).Value = id_maquina;
        //        cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker3.Value.Day.ToString() + "/" + dateTimePicker3.Value.Month.ToString() + "/" + dateTimePicker3.Value.Year.ToString();
        //        cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtCostoMant.Text;
        //        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescMant.Text;
        //        try
        //        {
        //            cmd.ExecuteNonQuery();
        //            MessageBox.Show("Mantenimiento agregado.");
        //            conn.Close();
        //        }
        //        catch (OleDbException ex)
        //        {
        //            MessageBox.Show(ex.Source);
        //            conn.Close();
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Connection Failed");
        //    }
        //}

        public void agregarReparacion()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO MaquinaReparacion (id_Maquina,Fecha,Costo,Descripcion) VALUES (@id,@Fecha,@Costo,@Descripcion)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@id", OleDbType.VarChar).Value = id_maquina;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
                cmd.Parameters.Add("@Costo", OleDbType.VarChar).Value = txtcostoRep.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescRep.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Reparacion agregada.");
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

        public void totalReparacion(DataGridView data)
        {
            int total = 0;
            foreach (DataGridViewRow row in data.Rows)
            {
                total += Int32.Parse(row.Cells[4].Value.ToString());
            }
            label3.Text = "Total: " + String.Format("{0:c}", total);
        }

        public frmVerMaquinaria(string id)
        {
            id_maquina = id;
            InitializeComponent();
            //cargarMaquinaAceite();
            cargarMaquinaReparacion();
            //cargarMaquinaMantenimiento();
            cargarMaquinaria();
            //totalMantenimiento();
            //totalAceite();
            totalReparacion();
            cargarOrdenes(id, DateTime.Now.Year.ToString());
            cargarOrdenes(id, (DateTime.Now.Year-1).ToString());
            cargarOrdenes(id);
            cargarInsumos();
            comboBox1.SelectedIndex = -1;
            totalReparacion(dataGridView6);
        }


        public void cargarOrdenes(string id)
        {
            string query = "SELECT h.ID, a.Actividad, h.fechaInicio, (t.Nombres +  ' ' + t.Apellidos) AS Supervisor , h.Costo FROM Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN (historicoOrdenes AS h INNER JOIN ordenMaquinas AS m ON h.ID = m.Orden) ON t.ID = h.Supervisor) ON a.ID = h.Actividad WHERE h.estadoOrden = 'Cerrada' AND m.Maquina =" +  id + " AND a.Actividad LIKE '%Mantenimiento " + dataGridView3.Rows[0].Cells[3].Value.ToString() + "%'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView2.DataSource = banco;
            dataGridView2.Columns[0].HeaderText = "Orden de Trabajo #";
            dataGridView2.Columns[2].HeaderText = "Fecha de Inicio";
        }
        public void cargarOrdenes(string id,string año)
        {
            string query = "SELECT h.ID, a.Actividad, h.fechaInicio, (t.Nombres +  ' ' + t.Apellidos) AS Supervisor , h.Costo FROM Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN (historicoOrdenes AS h INNER JOIN ordenMaquinas AS m ON h.ID = m.Orden) ON t.ID = h.Supervisor) ON a.ID = h.Actividad WHERE h.fechaInicio LIKE '%" + año + "%' AND h.estadoOrden = 'Cerrada' AND m.Maquina =" + id + " AND a.Actividad LIKE '%Mantenimiento " + dataGridView3.Rows[0].Cells[3].Value.ToString() + "%'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            if (año.Equals(DateTime.Now.Year.ToString()))
            {
                dataGridView6.DataSource = banco;
                dataGridView6.Columns[0].HeaderText = "Orden de Trabajo #";
                dataGridView6.Columns[2].HeaderText = "Fecha de Inicio";
            }
            else
            {
                dataGridView4.DataSource = banco;
                dataGridView4.Columns[0].HeaderText = "Orden de Trabajo #";
                dataGridView4.Columns[2].HeaderText = "Fecha de Inicio";
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAceite_Click(object sender, EventArgs e)
        {
            //agregarAceite();
            //cargarMaquinaAceite();
            //totalAceite();
        }

        public int getMaxID()
        {
            string query = "SELECT MAX(id) FROM maquinaReparacion";
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

        public void gastInsumos()
        {
            for (int i = 0; i < insumos.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE insumos SET Cantidad_Stock = Cantidad_Stock - @Cantidad WHERE ID = " + insumos[i]);
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = cantidad[i];
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

        private void btnReparacion_Click(object sender, EventArgs e)
        {
            agregarReparacion();
            cargarMaquinaReparacion();
            totalReparacion();
            int id = getMaxID();
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                agregarDetalles(id, dataGridView5.Rows[i].Cells[1].Value.ToString(), dataGridView5.Rows[i].Cells[0].Value.ToString());
            }
            gastInsumos();
        }

        public void eliminarAceite()
        {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el cambio de aceite de la fecha " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString()+"?", "Confirmar", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {

                    string id = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("DELETE FROM MaquinaAceite WHERE id = " + id);
                    cmd.Connection = conn;
                    conn.Open();

                    if (conn.State == ConnectionState.Open)
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Cambio de aceite eliminado.");
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

        public void eliminarReparacion()
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la reparacion de la fecha " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM MaquinaReparacion WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Reparacion eliminada.");
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

        //public void eliminarMantenimiento()
        //{
        //    DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la reparacion de la fecha " + dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[2].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

        //    if (dialogResult == DialogResult.Yes)
        //    {

        //        string id = dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString();
        //        conn.ConnectionString = connectionString;
        //        OleDbCommand cmd = new OleDbCommand("DELETE FROM MaquinaMantenimiento WHERE id = " + id);
        //        cmd.Connection = conn;
        //        conn.Open();

        //        if (conn.State == ConnectionState.Open)
        //        {
        //            try
        //            {
        //                cmd.ExecuteNonQuery();
        //                MessageBox.Show("Mantenimiento eliminado.");
        //                conn.Close();
        //                cargarMaquinaria();
        //            }
        //            catch (OleDbException ex)
        //            {
        //                MessageBox.Show(ex.Source);
        //                conn.Close();
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("Connection Failed");
        //        }
        //    }
        //}

        private void btnEliminarAceite_Click(object sender, EventArgs e)
        {
            //eliminarAceite();
            //cargarMaquinaAceite();
            //totalAceite();
        }

        private void btnEliminarRep_Click(object sender, EventArgs e)
        {
            eliminarReparacion();
            cargarMaquinaReparacion();
            totalReparacion();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //agregarMantenimiento();
            //cargarMaquinaMantenimiento();
            //totalMantenimiento();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //eliminarMantenimiento();
            //cargarMaquinaMantenimiento();
            //totalMantenimiento();
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals("") && !textBox2.Text.Equals(""))
            {
                dataGridView5.Rows.Add();                
                dataGridView5.Rows[dataGridView5.Rows.Count-1].Cells[0].Value = textBox2.Text;
                dataGridView5.Rows[dataGridView5.Rows.Count-1].Cells[1].Value = textBox1.Text;
                int costo = 0;
                for (int i = 0; i < dataGridView5.Rows.Count; i++)
                {
                    costo += Int32.Parse(dataGridView5.Rows[i].Cells[0].Value.ToString());
                }
                txtcostoRep.Text = costo.ToString();
                if (!comboBox1.Text.Equals(""))
                {
                    insumos.Add(comboBox1.SelectedValue.ToString());
                    cantidad.Add(textBox3.Text);
                }
            }
            else
                MessageBox.Show("Favor llenar el costo y el detalle.");
            comboBox1.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "1";
            textBox4.Text = "";
        }

        public void cargarInsumos()
        {
            string query = "SELECT ID,(Codigo + ' ' + Marca + ' ' + Modelo) As Insumo FROM Insumos";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable departamentos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.DisplayMember = "Insumo";
            comboBox1.ValueMember = "ID";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView5.Rows.Remove(dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex]);
        }

        public void cargarDetalles(string id)
        {
            insumos.Clear();
            cantidad.Clear();
            while (dataGridView5.Rows.Count != 0)
            {
                dataGridView5.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM reparacionDetalles WHERE Reparacion = " + id;
            //Ejecutar el query y llenar el GridView.
            try
            {
                conn.ConnectionString = connectionString;
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataReader reader = cmd.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    dataGridView5.Rows.Add();
                    dataGridView5.Rows[i].Cells[1].Value = reader.GetString(2);
                    dataGridView5.Rows[i].Cells[0].Value = reader.GetInt32(3);
                    i++;
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cargarDetalles(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
        }

        private void frmVerMaquinaria_Load(object sender, EventArgs e)
        {

        }

        public void getInsumo(string id)
        {
            string query = "SELECT ID,(Codigo + ' ' + Marca + ' ' + Modelo),Costo_Unitario FROM Insumos WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            try
            {
                conn.ConnectionString = connectionString;
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    textBox1.Text = reader.GetString(1);
                    textBox2.Text = reader.GetInt32(2).ToString();
                    textBox4.Text = reader.GetInt32(2).ToString();
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && !comboBox1.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                getInsumo(comboBox1.SelectedValue.ToString());
                textBox3.Text = "1";
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (!textBox3.Text.Equals("") && !textBox4.Text.Equals(""))
                textBox2.Text = (Double.Parse(textBox4.Text) * Double.Parse(textBox3.Text)).ToString();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals(""))
            {
                textBox4.Text = "";
                textBox3.Text = "1";
                textBox2.Text = "";
                textBox1.Text = "";
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString(), 1);
            newFrm.Show();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                totalReparacion(dataGridView6);
            }
            if (tabControl1.SelectedIndex == 1)
            {
                totalReparacion(dataGridView4);
            }
            if (tabControl1.SelectedIndex == 2)
            {
                totalReparacion(dataGridView2);
            }
        }

    }
}
