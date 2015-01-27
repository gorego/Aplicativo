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
    public partial class frmUsuarios : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarUsuarios()
        {
            string query = "SELECT d.ID,d.usuario,d.contrasena,s.ID, (s.Nombres + ' ' + s.Apellidos) AS Empleado, d.Tipo FROM Usuarios AS d INNER JOIN Trabajadores AS s ON d.Trabajador = s.ID";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            gridSupervisor.DataSource = supervisores;
            gridSupervisor.Columns[0].Visible = false;
            gridSupervisor.Columns[3].Visible = false;
            gridSupervisor.Columns[2].Visible = false;
            //for (int i = 0; i < gridSupervisor.Rows.Count; i++)
            //{
            //    if (gridSupervisor.Rows[i].Cells[5].Value.Equals("1"))
            //        gridSupervisor.Rows[i].Cells[5].Value = "Administrador";
            //    else
            //        gridSupervisor.Rows[i].Cells[5].Value = "Supervisor";
            //}
        }

        public void cargarOrdenes(int id)
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT ID From historicoOrdenes WHERE estadoOrden = 'Activa' AND Supervisor = " + id;
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
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = myReader.GetInt32(0);
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

        public void cargarEmpleados()
        {
            string query = "SELECT ID, (Nombres + ' ' + Apellidos) As Empleado FROM Trabajadores";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable departamentos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtEmpleado.DataSource = ds.Tables[0];
            txtEmpleado.DisplayMember = "Empleado";
            txtEmpleado.ValueMember = "ID";
            txtEmpleado.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtEmpleado.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void agregarUsuario()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Usuarios (usuario,contrasena,tipo,Trabajador) VALUES (@usuario,@contrasena,@tipo,@Trabajador)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@usuario", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@contrasena", OleDbType.VarChar).Value = textBox2.Text;
                if (radioButton1.Checked == true)
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 1;
                else if(radioButton2.Checked == true)
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 2;
                else if (radioButton3.Checked == true)
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 3;
                else
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 4;
                cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = txtEmpleado.SelectedValue;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Usuario agregado.");
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

        public frmUsuarios()
        {
            InitializeComponent();
            cargarUsuarios();
            cargarEmpleados();
            txtEmpleado.SelectedIndex = -1;
        }

        public void agregarLog(string accion, string usuario)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoIngresos (Usuario,Accion,Fecha) VALUES (@Usuario,@Accion,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Usuario", OleDbType.VarChar).Value = usuario;
                cmd.Parameters.Add("@Accion", OleDbType.VarChar).Value = accion;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year + " - " + DateTime.Now.Hour + ":" + DateTime.Now.ToString("mm") + ":" + DateTime.Now.ToString("ss");
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

        public string getUsuario(string id)
        {
            string usuario = "";
            string query = "SELECT usuario FROM Usuarios WHERE trabajador = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();            
            try
            {
                if (myReader.Read())
                {
                    usuario = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return usuario;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                if (!textBox2.Text.Equals(""))
                {
                    if (!txtEmpleado.Text.Equals(""))
                    {
                        agregarUsuario();
                        cargarUsuarios();
                        agregarLog("Usuario ha sido agregado.", textBox1.Text);
                        textBox1.Text = "";
                        textBox2.Text = "";
                        txtEmpleado.Text = "";
                        radioButton2.Checked = true;
                    }
                    else
                    {
                        MessageBox.Show("Favor seleccionar un empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("Favor ingresar la contraseña.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Favor ingresar el usuario.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void gridSupervisor_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtEmpleado.SelectedValue = gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[3].Value;
            cargarOrdenes(Int32.Parse(gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[3].Value.ToString()));
            if (gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[5].Value.ToString().Equals("1"))
                radioButton1.Checked = true;
            else
                radioButton2.Checked = true;
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el usuario " + gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Usuarios WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Usuario eliminado.");
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
                cargarUsuarios();
            }
        }

        public void modificarUsuario()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Usuarios SET usuario=@usuario,contrasena=@contrasena,tipo=@tipo,trabajador=@trabajador WHERE ID = " + gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@usuario", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@contrasena", OleDbType.VarChar).Value = textBox2.Text;
                if(radioButton1.Checked)
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 1;
                else if(radioButton2.Checked)
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 2;
                else if(radioButton3.Checked)
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 3;
                else
                    cmd.Parameters.Add("@tipo", OleDbType.VarChar).Value = 4;
                cmd.Parameters.Add("@trabajador", OleDbType.VarChar).Value = txtEmpleado.SelectedValue;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Usuario modficiado.");
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
            modificarUsuario();
            cargarUsuarios();
            textBox1.Text = "";
            textBox2.Text = "";
            txtEmpleado.Text = "";
            radioButton2.Checked = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            txtEmpleado.Text = "";
            radioButton2.Checked = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            darPermiso();
            agregarLog("Usuario otorgado permisos de modificación de asistencia de fijos", textBox1.Text);
            MessageBox.Show("Permiso otorgado.");
        }

        public void darPermiso()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE fijoSemanal SET Editable=1 WHERE Supervisor = " + txtEmpleado.SelectedValue + " AND Semana = @semana AND Ano = @ano");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                DateTime date1 = DateTime.Now;
                Calendar cal = dfi.Calendar;
                int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek) - 1;
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@ano", OleDbType.VarChar).Value = DateTime.Now.Year;
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

        public int getSemana(int id) 
        { 
            int semana = 0;
            string date ="";
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            IFormatProvider culture = new System.Globalization.CultureInfo("es-CO", true);        
            string query = "SELECT fechaInicio FROM historicoOrdenes WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();            
            try
            {
                if (myReader.Read())
                {

                    date = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            DateTime date1 = DateTime.Parse(date, culture, System.Globalization.DateTimeStyles.AssumeLocal);
            Calendar cal = dfi.Calendar;
            semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            return semana;
        }

        public void darPermiso(int orden, int semana)
        {            
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Control set Editable = 1 WHERE Orden = " + orden + " AND Semana = " + semana);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Permisos otorgados.");
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

        private void button5_Click(object sender, EventArgs e)
        {
            DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;            
            Calendar cal = dfi.Calendar;
            DateTime date1 = DateTime.Now;            
            int semanaActual = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);  
            int semanaOrden;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {                                    
                ch1 = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[1];
                if ((bool)ch1.FormattedValue == true) 
                {
                    semanaOrden = getSemana(Int32.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString()));
                    darPermiso(Int32.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString()),semanaActual - semanaOrden -1);
                    agregarLog("Usuario otorgado permisos de modificación en la orden de trabajo #" + dataGridView1.Rows[i].Cells[0].Value.ToString(),textBox1.Text);                    
                }
                
            }            
        }

    }
}
