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
    public partial class frmPropietarios : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarPropietarios()
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM Propietarios WHERE Nombre <> 'N/A'";
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
                    dataGridView1.Rows[i].Cells[1].Value = myReader.GetValue(1);
                    dataGridView1.Rows[i].Cells[2].Value = myReader.GetValue(2);
                    dataGridView1.Rows[i].Cells[3].Value = myReader.GetValue(3);
                    dataGridView1.Rows[i].Cells[4].Value = myReader.GetValue(4);
                    dataGridView1.Rows[i].Cells[5].Value = myReader.GetValue(6);
                    dataGridView1.Rows[i].Cells[6].Value = myReader.GetValue(5);
                    if (!myReader.GetValue(5).ToString().Equals(""))
                    {
                        dataGridView1.Rows[i].Cells[7].Value = Int32.Parse(DateTime.Now.Year.ToString()) - Int32.Parse(myReader.GetValue(5).ToString());
                    }
                    
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
            dataGridView1.Columns[0].Visible = false;
        }

        public void modificarPropietario()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Propietarios SET Nombre=@Nombre, Cedula=@Cedula, Direccion=@Direccion, Telefono=@Telefono, Correo=@Correo, fechaNacimiento=@fechaNacimiento WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Direccion", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Telefono", OleDbType.VarChar).Value = textBox4.Text;
                if (textBox5.Text.Equals(""))
                    cmd.Parameters.Add("@Correo", OleDbType.VarChar).Value = "NA";
                else
                    cmd.Parameters.Add("@Correo", OleDbType.VarChar).Value = textBox5.Text;
                cmd.Parameters.Add("@fechaNacimiento", OleDbType.VarChar).Value = textBox6.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Propietario modificado.");
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

        public void cargarPredios() {
            while (dataGridView2.Rows.Count != 0)
            {
                dataGridView2.Rows.RemoveAt(0);
            }
            string query = "SELECT b.ID, b.Codigo_predio, b.Predio FROM Propietarios AS p INNER JOIN BancoTierras AS b ON p.ID = b.Propietario WHERE p.id = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
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
                    dataGridView2.Rows.Add();
                    dataGridView2.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView2.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView2.Rows[i].Cells[1].Value = myReader.GetString(1);
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

        public void agregarPropietario()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Propietarios (Nombre,Cedula,Direccion, Telefono,Correo,fechaNacimiento) VALUES (@Nombre,@Cedula,@Direccion,@Telefono,@Correo,@fechaNacimiento)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Nombre", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Direccion", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Telefono", OleDbType.VarChar).Value = textBox4.Text;
                if (textBox5.Text.Equals(""))
                    cmd.Parameters.Add("@Correo", OleDbType.VarChar).Value = "NA";
                else
                    cmd.Parameters.Add("@Correo", OleDbType.VarChar).Value = textBox5.Text; 
                cmd.Parameters.Add("@fechaNacimiento", OleDbType.VarChar).Value = textBox6.Text; 
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Propietario agregado.");
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

        public void reiniciarTablero() {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
        }

        public frmPropietarios()
        {
            InitializeComponent();
            cargarPropietarios();
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                agregarPropietario();
                cargarPropietarios();
                reiniciarTablero();
            }
            else
                MessageBox.Show("Favor ingresar nombre del propietario", "Error");
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            if (!textBox1.Text.Equals(""))
            {
                modificarPropietario();
                cargarPropietarios();
                reiniciarTablero();
            }
            else
                MessageBox.Show("Favor ingresar nombre del propietario", "Error");
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar la unidad " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Propietarios WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Propietario eliminado.");
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
                cargarPropietarios();
                reiniciarTablero();
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            cargarPredios();
        }

        private void btnLotes_Click(object sender, EventArgs e)
        {
            frmBanco newFrm = new frmBanco();
            newFrm.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Variables.imprimir(dataGridView1);
        }
    }
}
