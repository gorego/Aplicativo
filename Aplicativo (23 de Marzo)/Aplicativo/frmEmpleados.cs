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
    public partial class frmEmpleados : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public void cargarEmpleados() {
            string query = "SELECT t.ID, t.Nombres, t.Apellidos, t.Cedula, t.fechaNacimiento, t.grupoSanguineo, t.Celular, t.Direccion, d.Departamento, c.Cargo, (m.Tipo+' / '+m.Marca +' / ' + m.Placa) As Maquinaria,t.fechaIngreso,t.fechaTerminacion, t.Pantalon, t.Camisa, t.Botas, t.Tipo, t.Ubicacion,t.diasLaborados FROM Maquinarias AS m INNER JOIN (Departamentos AS d INNER JOIN (CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) ON d.ID = t.Departamento) ON m.ID = t.Maquina";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable empleados = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(empleados);
            dataGridView1.DataSource = empleados;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].HeaderText = "Fecha de Nacimiento";
            dataGridView1.Columns[5].HeaderText = "G. Sanguinieo";
            dataGridView1.Columns[11].HeaderText = "Fecha de Ingreso";
            dataGridView1.Columns[12].HeaderText = "Fecha de Terminación";
            dataGridView1.Columns[18].HeaderText = "Dias Laborados";
            dataGridView1.Columns[16].Visible = false;
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
            txtDepartamento.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtDepartamento.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void modificarEmpleado() {
            if (!Todos.Checked)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores SET Nombres=@Nombres,Apellidos=@Apellidos,Cedula=@Cedula,fechaNacimiento=@fechaNacimiento,grupoSanguineo=@grupoSanguineo,Celular=@Celular,Direccion=@Direccion,Departamento=@Departamento,Cargo=@Cargo,Maquina=@Maquina,Pantalon=@Pantalon,Camisa=@Camisa,Botas=@Botas,Tipo=@Tipo,fechaIngreso=@fechaIngreso,fechaTerminacion=@fechaTerminacion,Ubicacion=@Ubicacion,diasLaborados=@diasLaborados WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Nombres", OleDbType.VarChar).Value = txtNombre.Text;
                    cmd.Parameters.Add("@Apellidos", OleDbType.VarChar).Value = txtApellido.Text;
                    cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = txtCedula.Text;
                    cmd.Parameters.Add("@fechaNacimiento", OleDbType.VarChar).Value = dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
                    cmd.Parameters.Add("@grupoSanguineo", OleDbType.VarChar).Value = txtSangre.Text;
                    cmd.Parameters.Add("@Celular", OleDbType.VarChar).Value = txtCelular.Text;
                    cmd.Parameters.Add("@Direccion", OleDbType.VarChar).Value = txtDir.Text;
                    cmd.Parameters.Add("@Departamento", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
                    cmd.Parameters.Add("@Cargo", OleDbType.VarChar).Value = txtCargo.SelectedValue;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = txtMaquinaria.SelectedValue;
                    cmd.Parameters.Add("@Pantalon", OleDbType.VarChar).Value = txtPantalon.Text;
                    cmd.Parameters.Add("@Camisa", OleDbType.VarChar).Value = txtCamisa.Text;
                    cmd.Parameters.Add("@Botas", OleDbType.VarChar).Value = txtBotas.Text;
                    if (fijoTH.Checked)
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "TH";
                    else if (fijoFHCTH.Checked)
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "FCTH";
                    else
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "Contratista";
                    cmd.Parameters.Add("@fechaIngreso", OleDbType.VarChar).Value = dateTimePicker2.Value.Day.ToString() + "/" + dateTimePicker2.Value.Month.ToString() + "/" + dateTimePicker2.Value.Year.ToString();
                    if (textTerminado.Text.Equals("Si"))
                        cmd.Parameters.Add("@fechaTerminacion", OleDbType.VarChar).Value = dateTimePicker3.Value.Day.ToString() + "/" + dateTimePicker3.Value.Month.ToString() + "/" + dateTimePicker3.Value.Year.ToString();
                    else
                        cmd.Parameters.Add("@fechaTerminacion", OleDbType.VarChar).Value = "";
                    cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;
                    cmd.Parameters.Add("@diasLaborados", OleDbType.VarChar).Value = textBox2.Text;
                    try
                    {
                        subirArchivo(textBox12);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Empleado modificado.");
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
            else 
            {
                MessageBox.Show("Favor seleccionar el tipo de empleado.", "Error");
            }
        }

        public void cargarCargos()
        {
            string query = "SELECT * FROM CargoLaboral";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable cargos = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtCargo.DataSource = ds.Tables[0];
            txtCargo.DisplayMember = "Cargo";
            txtCargo.ValueMember = "ID";
            txtCargo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtCargo.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarMaquinaria()
        {
            string query = "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtMaquinaria.DataSource = ds.Tables[0];
            txtMaquinaria.DisplayMember = "Maquina";
            txtMaquinaria.ValueMember = "ID";
            txtMaquinaria.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtMaquinaria.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void reiniciarTablero() {
            txtNombre.Text = "";
            txtApellido.Text = "";
            txtCedula.Text = "";
            textBox1.Text = "";
            textBox2.Text = "0";
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            txtSangre.Text = "";
            txtCelular.Text = "";
            txtDir.Text = "";
            txtDepartamento.Text = "";
            txtCargo.Text = "";
            txtMaquinaria.Text = "";
            txtPantalon.Text = "";
            txtCamisa.Text = "";
            txtBotas.Text = "";
            textBox12.Text = "";
            textTerminado.Text = "No";
            Todos.Checked = true;
        }

        public void agregarEmpleado()
        {
            if(!Todos.Checked){
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO Trabajadores(Nombres,Apellidos,Cedula,fechaNacimiento,grupoSanguineo,Celular,Direccion,Departamento,Cargo,Maquina,Pantalon,Camisa,Botas,Tipo,fechaIngreso,fechaTerminacion,Ubicacion,diasLaborados) VALUES (@Nombres,@Apellidos,@Cedula,@fechaNacimiento,@grupoSanguineo,@Celular,@Direccion,@Departamento,@Cargo,@Maquina,@Pantalon,@Camisa,@Botas,Tipo,@fechaIngreso,@fechaTerminacion,@Ubicacion,@diasLaborados)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Nombres", OleDbType.VarChar).Value = txtNombre.Text;
                    cmd.Parameters.Add("@Apellidos", OleDbType.VarChar).Value = txtApellido.Text;
                    cmd.Parameters.Add("@Cedula", OleDbType.VarChar).Value = txtCedula.Text;
                    cmd.Parameters.Add("@fechaNacimiento", OleDbType.VarChar).Value = dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
                    cmd.Parameters.Add("@grupoSanguineo", OleDbType.VarChar).Value = txtSangre.Text;
                    cmd.Parameters.Add("@Celular", OleDbType.VarChar).Value = txtCelular.Text;
                    cmd.Parameters.Add("@Direccion", OleDbType.VarChar).Value = txtDir.Text;
                    cmd.Parameters.Add("@Departamento", OleDbType.VarChar).Value = txtDepartamento.SelectedValue;
                    cmd.Parameters.Add("@Cargo", OleDbType.VarChar).Value = txtCargo.SelectedValue;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = txtMaquinaria.SelectedValue;
                    cmd.Parameters.Add("@Pantalon", OleDbType.VarChar).Value = txtPantalon.Text;
                    cmd.Parameters.Add("@Camisa", OleDbType.VarChar).Value = txtCamisa.Text;
                    cmd.Parameters.Add("@Botas", OleDbType.VarChar).Value = txtBotas.Text;
                    if (fijoTH.Checked)
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "TH";
                    else if (fijoFHCTH.Checked)
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "FCTH";
                    else
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = "Contratista";
                    cmd.Parameters.Add("@fechaIngreso", OleDbType.VarChar).Value = dateTimePicker2.Value.Day.ToString() + "/" + dateTimePicker2.Value.Month.ToString() + "/" + dateTimePicker2.Value.Year.ToString();
                    if (textTerminado.Text.Equals("Si"))
                        cmd.Parameters.Add("@fechaTerminacion", OleDbType.VarChar).Value = dateTimePicker3.Value.Day.ToString() + "/" + dateTimePicker3.Value.Month.ToString() + "/" + dateTimePicker3.Value.Year.ToString();
                    else
                        cmd.Parameters.Add("@fechaTerminacion", OleDbType.VarChar).Value = "";
                    cmd.Parameters.Add("@Ubicacion", OleDbType.VarChar).Value = textBox1.Text;
                    cmd.Parameters.Add("@diasLaborados", OleDbType.VarChar).Value = textBox2.Text;
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Empleado agregado.");
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
            else
            {
                MessageBox.Show("Favor seleccionar el tipo de empleado.", "Error");
            }
        }

        public frmEmpleados()
        {
            InitializeComponent();
            cargarEmpleados();
            cargarCargos();
            cargarDepartamentos();
            cargarMaquinaria();
            txtMaquinaria.SelectedItem = null;
            txtDepartamento.SelectedItem = null;
            txtCargo.SelectedItem = null;
            textTerminado.Text = "No";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd/MM/yyyy";
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView1.Columns[2].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void btnCargos_Click(object sender, EventArgs e)
        {
            frmCargosLaborales newFrm = new frmCargosLaborales();
            newFrm.Show();
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!txtNombre.Text.Equals(""))
            {
                if (!txtApellido.Text.Equals(""))
                {
                    if (!txtCedula.Text.Equals(""))
                    {
                        if (!txtDepartamento.Text.Equals(""))
                        {
                            if (!txtCargo.Text.Equals(""))
                            {
                                if (!txtMaquinaria.Text.Equals(""))
                                {
                                    agregarEmpleado();
                                    cargarEmpleados();
                                }
                                else
                                {
                                    MessageBox.Show("Favor ingresar maquinaria asociada al empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Favor ingresar el cargo del empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Favor ingresar departamento del empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Favor ingresar cedula del empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }            
                }
                else
                {
                    MessageBox.Show("Favor ingresar apellido del empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Favor ingresar nombre del empleado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            modificarEmpleado();
            cargarEmpleados();
            reiniciarTablero();
        }

        public void subirArchivo(TextBox text)
        {
            string arch = text.Text;
            if (!arch.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text);
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text, txtNombre.Text + " " + txtApellido.Text + "*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(arch, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text);
                    string ext = Path.GetExtension(arch);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text + "\\" + txtNombre.Text + " " + txtApellido.Text + ext));
                }
            }
        }


        public void cargarImagen()
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text);
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text, txtNombre.Text + " " + txtApellido.Text + "*");
            if (prueba.Length > 0)
            {
                if (File.Exists(prueba[0]))
                {
                    pictureBox1.Image = Image.FromFile(prueba[0]);                    
                }
            }
            else
            {
                pictureBox1.Image = pictureBox1.ErrorImage;
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1 || dataGridView1.CurrentCell.ColumnIndex == 2)
            {
                frmLiquidacion newFrm = new frmLiquidacion(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
            txtNombre.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtApellido.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtCedula.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString();            
            dateTimePicker1.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString());
            txtSangre.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            txtCelular.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtDir.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            txtDepartamento.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            txtCargo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            txtMaquinaria.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            if (!dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString().Equals(""))
            {
                dateTimePicker2.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString());
            }
            if (!dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString().Equals(""))
            {
                dateTimePicker3.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString());
                textTerminado.Text = "Si";
            }
            else
            {
                textTerminado.Text = "No";
            }
            txtPantalon.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            txtCamisa.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            txtBotas.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString().Equals("TH"))
                fijoTH.Checked = true;
            else if ((dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString().Equals("FCTH")))
                fijoFHCTH.Checked = true;
            else
                Contratista.Checked = true;
            cargarImagen();
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar a " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + " " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString()+"?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM Trabajadores WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Empleado eliminado.");
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
                cargarEmpleados();
                reiniciarTablero();
            }
        }

        public void buscarEmpleado()
        {
            string query = "SELECT t.ID, t.Nombres, t.Apellidos, t.Cedula, t.fechaNacimiento, t.grupoSanguineo, t.Celular, t.diasLaborados, t.Direccion, d.Departamento, c.Cargo, (m.Tipo+' / '+m.Marca +' / ' + m.Placa) As Maquinaria, t.fechaIngreso, t.fechaTerminacion, t.Pantalon, t.Camisa, t.Botas, t.Tipo FROM Maquinarias AS m INNER JOIN (Departamentos AS d INNER JOIN (CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) ON d.ID = t.Departamento) ON m.ID = t.Maquina ";
            int i = 0;
            if (!txtNombre.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Nombres LIKE '%" + txtNombre.Text + "%'";
            }
            if (!txtApellido.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Apellidos LIKE '%" + txtApellido.Text + "%'";
            }
            if (!txtCedula.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Cedula LIKE '%" + txtCedula.Text + "%'";
            }
            string date1 = dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Year.ToString();
            string date2 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString();
            if (!date1.Equals(date2))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.fechaNacimiento LIKE '%" + date1 + "%'";
            }
            date1 = dateTimePicker2.Value.Month.ToString() + "/" + dateTimePicker2.Value.Day.ToString() + "/" + dateTimePicker2.Value.Year.ToString();
            date2 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString();
            if (!date1.Equals(date2))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.fechaIngreso LIKE '%" + date1 + "%'";
            }
            date1 = dateTimePicker3.Value.Month.ToString() + "/" + dateTimePicker3.Value.Day.ToString() + "/" + dateTimePicker3.Value.Year.ToString();
            date2 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString();
            if (textTerminado.Text.Equals("Si"))
            {
                if (!date1.Equals(date2))
                {
                    if (i != 0)
                        query += " AND ";
                    else
                        query += "WHERE ";
                    i++;
                    query += "t.fechaTerminacion LIKE '%" + date1 + "%'";
                }
            }
            if (!txtSangre.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.grupoSanguineo LIKE '%" + txtSangre.Text + "%'";
            }
            if (!txtCelular.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Celular LIKE '%" + txtCelular.Text + "%'";
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.diasLaborados LIKE '%" + textBox2.Text + "%'";
            }
            if (!txtDir.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Direccion LIKE '%" + txtDir.Text + "%'";
            }
            if (!txtDepartamento.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "d.Departamento LIKE '%" + txtDepartamento.Text + "%'";
            }
            if (!txtCargo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "c.Cargo LIKE '%" + txtCargo.Text + "%'";
            }
            if (!txtPantalon.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Pantalon LIKE '%" + txtPantalon.Text + "%'";
            }
            if (!txtCamisa.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Camisa LIKE '%" + txtCamisa.Text + "%'";
            }
            if (!txtBotas.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Botas LIKE '%" + txtBotas.Text + "%'";
            }
            if(fijoFHCTH.Checked){
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Tipo LIKE 'FCTH'";
            }
            else if (fijoTH.Checked) {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "t.Tipo LIKE 'TH'";
            }
            else if (Contratista.Checked)
            {
                    if (i != 0)
                        query += " AND ";
                    else
                        query += "WHERE ";
                    i++;
                    query += "t.Tipo LIKE 'Contratista'";
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

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            buscarEmpleado();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textTerminado_TextChanged(object sender, EventArgs e)
        {
            if (textTerminado.Text.Equals("Si"))
            {
                label15.Visible = true;
                dateTimePicker3.Visible = true;
            }
            else
            {
                label15.Visible = false;
                dateTimePicker3.Visible = false;
            }
        }

        public void imprimirEmpleados(DataGridView data)
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
                            excelCellrange.Interior.Color = System.Drawing.Color.LightGreen;
                            excelCellrange.AutoFilter(1);
                            //excelCellrange.Interior.Color = System.Drawing.Color.Blue;
                            //excelCellrange.Font.Color = System.Drawing.Color.White;
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
            imprimirEmpleados(dataGridView1);
        }

        private void btnPredial_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            textBox12.Text = openFileDialog1.FileName;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text);
            string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Anexos\\Empleados\\" + txtNombre.Text + " " + txtApellido.Text, txtNombre.Text + " " + txtApellido.Text + "*");
            if (prueba.Length > 0)
            {
                if (File.Exists(prueba[0]))
                {
                    System.Diagnostics.Process.Start(prueba[0]);
                }
            }
        }
    }
}
