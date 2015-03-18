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
    public partial class frmSupervisor : Form
    {
        //Base de datos.
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        //public void cargarSupervisores() {
        //    string query = "SELECT s.ID, s.Supervisor, s.CC, d.Departamento, s.Telefono FROM Departamentos AS d INNER JOIN Supervisores AS s ON d.ID = s.Departamento";
        //    //Ejecutar el query y llenar el GridView.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    DataTable supervisores = new DataTable();
        //    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        //    da.Fill(supervisores);
        //    gridSupervisor.DataSource = supervisores;
        //    gridSupervisor.Columns[0].Visible = false;
        //}

        public void cargarSupervisores()
        {
            string query = "SELECT s.ID, (s.Nombres + ' ' + s.Apellidos) AS Empleado, s.Cedula, d.Departamento, s.Celular FROM Departamentos AS d INNER JOIN Trabajadores AS s ON d.ID = s.Departamento WHERE s.Supervisor = 'Si'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            gridSupervisor.DataSource = supervisores;
            gridSupervisor.Columns[0].Visible = false;
        }


        //public void cargarDepartamentos() {
        //    string query = "SELECT * FROM Departamentos";
        //    //Ejecutar el query y llenar el ComboBox.
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand(query, conn);
        //    DataTable departamentos = new DataTable();
        //    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        //    DataSet ds = new DataSet();
        //    da.Fill(ds);
        //    txtDepartamento.DataSource = ds.Tables[0];
        //    txtDepartamento.DisplayMember = "Departamento";
        //    txtDepartamento.ValueMember = "ID";
        //    txtDepartamento.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        //    txtDepartamento.AutoCompleteSource = AutoCompleteSource.ListItems;
        //}

        public void agregarSupervisor() { 
            conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores SET Supervisor=@Supervisor WHERE ID = " + txtEmpleado.SelectedValue);
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = "Si";                    
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Supervisor agregado.");
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

        public int getDepartamento() { 
            string query = "SELECT Departamento FROM Trabajadores WHERE ID = "+txtEmpleado.SelectedValue+"";
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

        public void modificarDepartamento(int dep) {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Departamentos SET Supervisor=@Supervisor WHERE ID = " + dep);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = txtEmpleado.SelectedValue;
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

        //public void agregarSupervisor()
        //{
        //    conn.ConnectionString = connectionString;
        //    OleDbCommand cmd = new OleDbCommand("INSERT INTO Supervisores (Supervisor, CC, Telefono, Departamento) VALUES (@Supervisor,@CC,@Telefono,@Departamento)");
        //    cmd.Connection = conn;
        //    conn.Open();

        //    if (conn.State == ConnectionState.Open)
        //    {
        //        cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = txtSupervisor.Text;
        //        cmd.Parameters.Add("@CC", OleDbType.VarChar).Value = txtCC.Text;
        //        cmd.Parameters.Add("@Telefono", OleDbType.VarChar).Value = txtTelefono.Text;
        //        cmd.Parameters.Add("@Departamento", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
        //        try
        //        {
        //            cmd.ExecuteNonQuery();
        //            MessageBox.Show("Supervisor agregado.");
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

        //    txtCC.Text = "";
        //    txtSupervisor.Text = "";
        //    txtTelefono.Text = "";
        //    txtDepartamento.SelectedItem = null;
        //}

        public frmSupervisor()
        {
            InitializeComponent();
            cargarSupervisores();
            //cargarDepartamentos();
            cargarEmpleados();
            //txtDepartamento.SelectedItem = null;
            txtEmpleado.SelectedItem = null;
            gridSupervisor.Columns[1].DefaultCellStyle.Font = new Font(gridSupervisor.DefaultCellStyle.Font, FontStyle.Underline);
        }

        //private void btnAgregar_Click(object sender, EventArgs e)
        //{
        //    if (!txtSupervisor.Text.Equals("") && !txtCC.Text.Equals("") && !txtDepartamento.Text.Equals("") && !txtTelefono.Text.Equals(""))
        //    {
        //        agregarSupervisor();
        //        cargarSupervisores();
        //    }
        //    else
        //    {
        //        MessageBox.Show("Favor llenar todas las casillas.");
        //    }
        //}

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (gridSupervisor.CurrentCell.ColumnIndex == 1)
            {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar a " + gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.OKCancel);

                if (dialogResult == DialogResult.OK)
                {
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores SET Supervisor=@Supervisor WHERE ID = " + gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[0].Value.ToString());
                    cmd.Connection = conn;
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = "No";
                        try
                        {
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Supervisor eliminado.");
                            conn.Close();
                            cargarSupervisores();
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
                    //string cc = gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[2].Value.ToString();
                    //conn.ConnectionString = connectionString;
                    //OleDbCommand cmd = new OleDbCommand("DELETE FROM Supervisores WHERE CC = '" + cc + "'");
                    //cmd.Connection = conn;
                    //conn.Open();

                    //if (conn.State == ConnectionState.Open)
                    //{
                    //    try
                    //    {
                    //        cmd.ExecuteNonQuery();
                    //        MessageBox.Show("Supervisor eliminado.");
                    //        conn.Close();
                    //        cargarSupervisores();
                    //    }
                    //    catch (OleDbException ex)
                    //    {
                    //        MessageBox.Show(ex.Source);
                    //        conn.Close();
                    //    }
                    //}
                    //else
                    //{
                    //    MessageBox.Show("Connection Failed");
                    //}
                }
            }
            else
            {
                MessageBox.Show("Favor seleccionar nombre del supervisor.", "Error");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarSupervisor();
            cargarSupervisores();
            int dep = getDepartamento();
            modificarDepartamento(dep);
        }

        private void gridSupervisor_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (gridSupervisor.CurrentCell.ColumnIndex == 1)
            {
                frmSupervisorFijos newFrm = new frmSupervisorFijos(gridSupervisor.Rows[gridSupervisor.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }
    }
}
