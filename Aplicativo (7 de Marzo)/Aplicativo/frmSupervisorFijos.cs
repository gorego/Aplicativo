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
    public partial class frmSupervisorFijos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();        
        List<string> empleados = new List<string>();
        string fijos;

        public void cargarEmpleados()
        {
            string query = "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtEmpleado.DataSource = ds.Tables[0];
            txtEmpleado.DisplayMember = "nombre";
            txtEmpleado.ValueMember = "ID";
            txtEmpleado.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtEmpleado.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void getEmpleados(string id)
        {
            listBox2.Items.Clear();
            empleados.Clear();
            string query = "SELECT i.ID, i.Supervisor, (m.Nombres + ' ' + m.Apellidos) As nombre, m.ID FROM Trabajadores AS m INNER JOIN supervisorEmpleados AS i ON m.ID = i.Trabajador WHERE i.Supervisor = " + id;
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

        public void agregarSupervisorEmpleados(string id)
        {
            for (int i = 0; i < empleados.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO supervisorEmpleados(Supervisor,Trabajador) VALUES (@Supervisor,@Trabajador)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = id;
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = empleados[i];
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

        public void eliminarSupervisorEmpleados(string id)
        {
            for (int i = 0; i < empleados.Count+1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM supervisorEmpleados WHERE Supervisor = " + id);
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

        public void totalReparacion()
        {
            int total = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                total += Int32.Parse(dataGridView2.Rows[i].Cells[6].Value.ToString());
            }
            label1.Text = "Total: " + String.Format("{0:c}", total);
        }

        public frmSupervisorFijos(string id)
        {
            InitializeComponent();
            cargarEmpleados();
            getEmpleados(id);
            fijos = id;
            cargarOrdenes(id);
            totalReparacion();
        }

        public void cargarOrdenes(string id)
        {
            string query = "SELECT h.ID, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.Costo, h.estadoOrden FROM Areas AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h.Supervisor = " + id + " UNION ALL SELECT h.ID, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal, (t.Nombres + ' ' + t.Apellidos) As Supervisor, h.Costo, h.estadoOrden FROM LoteGanadero AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h.Supervisor = " + id + " UNION ALL SELECT h.ID, a.Actividad, area.Lote, h.fechaInicio, h.fechaFinal,(t.Nombres + ' ' + t.Apellidos) As Supervisor, h.Costo, h.estadoOrden FROM Lotes AS area INNER JOIN (Actividades AS a INNER JOIN (Trabajadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Supervisor) ON a.ID = h.Actividad) ON area.Codigo = h.Lote WHERE h.estadoOrden = 'Cerrada' AND h.Supervisor = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable banco = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(banco);
            dataGridView2.DataSource = banco;
            dataGridView2.Columns[0].HeaderText = "Orden de Trabajo #";
            dataGridView2.Columns[3].HeaderText = "Fecha de Inicio";
            dataGridView2.Columns[4].HeaderText = "Fecha de Finalización";
            dataGridView2.Columns[7].HeaderText = "Estado de la Orden";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
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

        private void button2_Click(object sender, EventArgs e)
        {
            eliminarSupervisorEmpleados(fijos);
            agregarSupervisorEmpleados(fijos);
            this.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString(), 1);
            newFrm.Show();
        }
    }
}
