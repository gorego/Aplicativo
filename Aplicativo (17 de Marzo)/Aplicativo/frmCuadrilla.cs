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
    public partial class frmCuadrilla : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        List<string> empleados = new List<string>();

        public frmCuadrilla()
        {
            InitializeComponent();
            Variables.cargar(txtEmpleado, "SELECT ID,(Nombres + ' ' + Apellidos) As Nombre FROM Trabajadores", "Nombre");
            Variables.cargar(dataGridView1, "SELECT * FROM Cuadrilla");
        }

        public void getEmpleados(string id)
        {
            listBox2.Items.Clear();
            empleados.Clear();
            string query = "SELECT i.ID, i.Cuadrilla, (m.Nombres + ' ' + m.Apellidos) As nombre, m.ID FROM Trabajadores AS m INNER JOIN empleadoCuadrilla AS i ON m.ID = i.Trabajador WHERE i.Cuadrilla = " + id;
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
                    listBox2.Items.Add(myReader.GetString(2));
                    empleados.Add(myReader.GetInt32(3).ToString());
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

        public void agregarCuadrilla()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Cuadrilla (Nombre) VALUES (@Nombre)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;           
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cuadrilla agregada.");
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

        public void eliminarCuadrilla()
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la cuadrilla " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {

                    string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                    conn.ConnectionString = connectionString;
                    OleDbCommand cmd = new OleDbCommand("DELETE FROM Cuadrilla WHERE id = " + id);
                    cmd.Connection = conn;
                    conn.Open();

                    if (conn.State == ConnectionState.Open)
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Cuadrilla eliminada.");
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
            else
            {
                MessageBox.Show("Favor seleccionar el nombre de la actividad.", "Error");
            }
        }


        public void modificarCuadrilla(string id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Cuadrilla SET Nombre=@Nombre WHERE ID = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cuadrilla modificada.");
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


        public void agregarEmpleados(int id)
        {
            for (int i = 0; i < empleados.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO empleadoCuadrilla(Trabajador,Cuadrilla) VALUES (@Trabajador,@Cuadrilla)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {                    
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = empleados[i];
                    cmd.Parameters.Add("@Cuadrilla", OleDbType.VarChar).Value = id;
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

        public int getMaxID()
        {
            string query = "SELECT MAX(id) FROM Cuadrilla";
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

        public void agregarEmpleadosCuadrilla(string id)
        {
            for (int i = 0; i < empleados.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO empleadoCuadrilla(Trabajador,Cuadrilla) VALUES (@Trabajador,@Cuadrilla)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {                    
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = empleados[i];
                    cmd.Parameters.Add("@Cuadrilla", OleDbType.VarChar).Value = id;
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

        public void eliminarEmpleadosCuadrilla(string id)
        {
            for (int i = 0; i < empleados.Count + 1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM empleadoCuadrilla WHERE Cuadrilla = " + id);
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

        private void button7_Click(object sender, EventArgs e)
        {
            listBox2.Items.Add(txtEmpleado.Text);
            empleados.Add(txtEmpleado.SelectedValue.ToString());    
        }

        private void button8_Click(object sender, EventArgs e)
        {
            empleados.RemoveAt(listBox2.SelectedIndex);
            listBox2.Items.Remove(listBox2.SelectedItem);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            getEmpleados(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            agregarCuadrilla();
            int act = getMaxID();
            agregarEmpleados(act);
            Variables.cargar(dataGridView1, "SELECT * FROM Cuadrilla");            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            modificarCuadrilla(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            eliminarEmpleadosCuadrilla(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            agregarEmpleadosCuadrilla(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            Variables.cargar(dataGridView1, "SELECT * FROM Cuadrilla");            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            eliminarEmpleadosCuadrilla(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            eliminarCuadrilla();
            Variables.cargar(dataGridView1, "SELECT * FROM Cuadrilla");            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
