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
    public partial class frmCargosLaborales : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        double salario = 0;

        public void cargarCargos() {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM CargoLaboral WHERE Cargo <> 'N/A'";
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
                    double salario = myReader.GetInt32(3);
                    double riesgo = double.Parse(myReader.GetString(5).Replace(".",string.Empty).Replace("%",string.Empty));
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView1.Rows[i].Cells[1].Value = myReader.GetString(1);
                    dataGridView1.Rows[i].Cells[2].Value = myReader.GetString(2);
                    dataGridView1.Rows[i].Cells[3].Value = String.Format("{0:c}", (salario));
                    dataGridView1.Rows[i].Cells[4].Value = String.Format("{0:c}", (salario*0.92));
                    dataGridView1.Rows[i].Cells[5].Value = myReader.GetString(4);
                    dataGridView1.Rows[i].Cells[6].Value = myReader.GetString(5)+"%";
                    dataGridView1.Rows[i].Cells[7].Value = String.Format("{0:c}", (salario/2));
                    dataGridView1.Rows[i].Cells[8].Value = String.Format("{0:c}", (salario/2));
                    dataGridView1.Rows[i].Cells[9].Value = String.Format("{0:c}", (salario));
                    dataGridView1.Rows[i].Cells[10].Value = String.Format("{0:c}", (salario*0.12));
                    dataGridView1.Rows[i].Cells[11].Value = String.Format("{0:c}", (salario/2));
                    double ILA = (salario * 12) + ((salario / 2) * 3) + (salario) + (salario * 0.12);
                    dataGridView1.Rows[i].Cells[12].Value = String.Format("{0:c}", (ILA));
                    dataGridView1.Rows[i].Cells[13].Value = String.Format("{0:c}", (salario*0.08));
                    dataGridView1.Rows[i].Cells[14].Value = String.Format("{0:c}", (salario*0.08));
                    dataGridView1.Rows[i].Cells[15].Value = String.Format("{0:c}", (salario*(double.Parse("0.0"+riesgo))));
                    dataGridView1.Rows[i].Cells[16].Value = String.Format("{0:c}", (salario*0.09));
                    double totalss = (salario * 0.08) + (salario * 0.08) + (salario * double.Parse("0.0" + riesgo)) + (salario * 0.09);
                    double total = (totalss*12)+ILA;
                    dataGridView1.Rows[i].Cells[17].Value = String.Format("{0:c}", totalss);
                    dataGridView1.Rows[i].Cells[18].Value = String.Format("{0:c}", total);
                    dataGridView1.Rows[i].Cells[19].Value = myReader.GetString(6);
                    dataGridView1.Rows[i].Cells[20].Value = myReader.GetString(7);
                    dataGridView1.Rows[i].Cells[21].Value = myReader.GetString(8);
                    dataGridView1.Rows[i].Cells[22].Value = myReader.GetString(9);
                    dataGridView1.Rows[i].Cells[23].Value = myReader.GetString(10);
                    dataGridView1.Rows[i].Cells[24].Value = myReader.GetString(11);
                    dataGridView1.Rows[i].Cells[25].Value = myReader.GetString(12);
                    dataGridView1.Rows[i].Cells[26].Value = myReader.GetString(13);
                    dataGridView1.Rows[i].Cells[27].Value = myReader.GetString(14);
                    dataGridView1.Rows[i].Cells[28].Value = myReader.GetString(15);
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

        public void agregarCargo() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO CargoLaboral (Cargo,Descripcion,Salario,Funciones,Riesgo,Casco,Tapa_Oidos,Careta,Lentes,Tapa_Boca,Delantales,Guantes,Pantalon_Anticorte,Bota_Anticorte,Botas_Puntera) VALUES (@Cargo,@Descripcion,@Salario,@Funciones,@Riesgo,@Casco,@Tapa_Oidos,@Careta,@Lentes,@Tapa_Boca,@Delantales,@Guantes,@Pantalon_Anticorte,@Bota_Anticrote,@Botas_Puntera)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cargo", OleDbType.VarChar).Value = txtNombre.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Salario", OleDbType.VarChar).Value = txtSalario.Text.Substring(1,txtSalario.TextLength-3);
                cmd.Parameters.Add("@Funciones", OleDbType.VarChar).Value = txtFunciones.Text;
                cmd.Parameters.Add("@Riesgo", OleDbType.VarChar).Value = txtRiesgo.Text.Substring(0, txtRiesgo.TextLength - 1);
                if(checkedListBox1.GetItemCheckState(0).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Casco", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Casco", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(1).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Tapa_Oidos", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Tapa_Oidos", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(2).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Careta", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Careta", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(3).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Lentes", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Lentes", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(4).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Tapa_Boca", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Tapa_Boca", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(5).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Delantales", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Delantales", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(6).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Guantes", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Guantes", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(7).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Pantalon_Anticorte", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Pantalon_Anticorte", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(8).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Botas_Anticorte", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Bota_Anticorte", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(9).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Bota_Puntera", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Botas_Puntera", OleDbType.VarChar).Value = "No";
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cargo agregada.");
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

        public void modificarCargo() {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE CargoLaboral SET Cargo=@Cargo,Descripcion=@Descripcion,Salario=@Salario,Funciones=@Funciones,Riesgo=@Riesgo,Casco=@Casco,Tapa_Oidos=@Tapa_Oidos,Careta=@Careta,Lentes=@Lentes,Tapa_Boca=@Tapa_Boca,Delantales=@Delantales,Guantes=@Guantes,Pantalon_Anticorte=@Pantalon_Anticorte,Bota_Anticorte=@Bota_Anticorte,Botas_Puntera=@Botas_Puntera WHERE ID = " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Cargo", OleDbType.VarChar).Value = txtNombre.Text;
                cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = txtDescripcion.Text;
                cmd.Parameters.Add("@Salario", OleDbType.VarChar).Value = txtSalario.Text.Substring(1, txtSalario.TextLength - 3);
                cmd.Parameters.Add("@Funciones", OleDbType.VarChar).Value = txtFunciones.Text;
                cmd.Parameters.Add("@Riesgo", OleDbType.VarChar).Value = txtRiesgo.Text.Substring(0, txtRiesgo.TextLength - 1);
                if (checkedListBox1.GetItemCheckState(0).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Casco", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Casco", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(1).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Tapa_Oidos", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Tapa_Oidos", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(2).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Careta", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Careta", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(3).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Lentes", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Lentes", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(4).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Tapa_Boca", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Tapa_Boca", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(5).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Delantales", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Delantales", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(6).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Guantes", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Guantes", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(7).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Pantalon_Anticorte", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Pantalon_Anticorte", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(8).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Botas_Anticorte", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Bota_Anticorte", OleDbType.VarChar).Value = "No";
                if (checkedListBox1.GetItemCheckState(9).ToString().Equals("Checked"))
                    cmd.Parameters.Add("@Bota_Puntera", OleDbType.VarChar).Value = "Si";
                else
                    cmd.Parameters.Add("@Botas_Puntera", OleDbType.VarChar).Value = "No";
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Cargo modificado.");
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

        public void eliminarCargo() {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el cargo laboral " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString()+"?", "Confirmar", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {

                string id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM CargoLaboral WHERE ID = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Cargo eliminado.");
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

        public void reiniciarTablero()
        {
            txtNombre.Text = "";
            txtFunciones.Text = "";
            txtDescripcion.Text = "";
            txtAnual.Text = "";
            txtRiesgo.Text = "";
            txtTotalSS.Text = "";
            txtSalarioEfect.Text = "";
            txtSalario.Text = "";
            txtPrimasJun.Text = "";
            txtPrimasDic.Text = "";
            txtCesantias.Text = "";
            txtIntereses.Text = "";
            txtVacaciones.Text = "";
            txtILA.Text = "";
            txtEps.Text = "";
            txtPension.Text = "";
            txtArl.Text = "";
            txtParafiscales.Text = "";
            foreach (int i in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }
        }

        public frmCargosLaborales()
        {
            InitializeComponent();
            cargarCargos();
            dataGridView1.Columns[2].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (!txtNombre.Text.Equals(""))
            {
                if (!txtRiesgo.Text.Equals(""))
                {
                    if (!txtSalario.Text.Equals(""))
                    {
                        agregarCargo();
                        subirArchivo(txtNombre.Text);
                        cargarCargos();
                        reiniciarTablero();
                    }
                    else
                        MessageBox.Show("Favor ingresar el salario del cargo.","Error");
                }
                else
                    MessageBox.Show("Favor ingresar el riesgo del cargo.","Error");
            }
            else
                MessageBox.Show("Favor ingresar nombre del cargo.", "Error");
        }

        public void CalculoSalario()
        {
            double ILA;
            txtSalarioEfect.Text = String.Format("{0:c}", (salario * (0.92)));
            txtSalario.Text = String.Format("{0:c}", (salario));
            txtPrimasJun.Text = String.Format("{0:c}", (salario / 2));
            txtPrimasDic.Text = String.Format("{0:c}", (salario / 2));
            txtCesantias.Text = String.Format("{0:c}", (salario));
            txtIntereses.Text = String.Format("{0:c}", (salario * 0.12));
            txtVacaciones.Text = String.Format("{0:c}", (salario / 2));
            ILA = (salario * 12) + ((salario / 2) * 3) + (salario) + (salario * 0.12);
            txtILA.Text = String.Format("{0:c}", ILA);
            txtEps.Text = String.Format("{0:c}", (salario * 0.08));
            txtPension.Text = String.Format("{0:c}", (salario * 0.08));
            double riesgo;
            if (txtRiesgo.Text.Trim().Equals(""))
                riesgo = 0;
            else
                riesgo = double.Parse(txtRiesgo.Text.Replace(".",string.Empty).Replace("%",string.Empty));
            txtArl.Text = String.Format("{0:c}", (salario * Double.Parse("0.0" + riesgo)));
            txtParafiscales.Text = String.Format("{0:c}", (salario * 0.09));
            double total = 0;
            double totalss = 0;
            totalss = (salario * 0.08) + (salario * 0.08) + (salario * double.Parse("0.0" + riesgo)) + (salario * 0.09);
            txtTotalSS.Text = String.Format("{0:c}", totalss);
            total = (totalss * 12) + ILA;
            txtAnual.Text = String.Format("{0:c}", total);
        }

        private void txtSalario_Leave(object sender, EventArgs e)
        {
            bool isNum = double.TryParse(txtSalario.Text.Trim(), out salario);
            if (isNum)
            {
                CalculoSalario();                
            }
            else
            {
                MessageBox.Show("Favor digitar un numero valido.", "Error");
            }
        }

        private void txtRiesgo_Leave(object sender, EventArgs e)
        {
            double riesgo;
            if (!txtRiesgo.Text.Contains('%'))
            {
                bool isNum = double.TryParse(txtRiesgo.Text.Trim('%'), out riesgo);
                if (isNum)
                {
                    CalculoSalario();
                    txtRiesgo.Text = riesgo + "%";
                }
                else
                {
                    MessageBox.Show("Favor digitar un numero valido.", "Error");
                }
            }
        }

        private void txtSalario_Enter(object sender, EventArgs e)
        {
            txtSalario.Text = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 2) {

                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\" + this.Text + "\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString());
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\" + this.Text + "\\" + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString(), "Descripcion*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {
                        System.Diagnostics.Process.Start(prueba[0]);
                    }
                    else
                    {
                        MessageBox.Show("No se encuentra el archivo.", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo.", "Error");
                }
            }
            salario = double.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString().Substring(1, dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString().Length - 4).Replace(".",string.Empty).Replace(",",string.Empty));
            txtSalario.Text = salario.ToString();
            txtRiesgo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            txtNombre.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            txtDescripcion.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString();
            txtFunciones.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            CalculoSalario();
            for (int i = 19; i < 29; i++)
            {
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[i].Value.ToString().Equals("Si"))
                    checkedListBox1.SetItemCheckState(i-19, CheckState.Checked);
                else
                    checkedListBox1.SetItemCheckState(i-19, CheckState.Unchecked);  
            }
        }

        private void btnReiniciar_Click(object sender, EventArgs e)
        {
            reiniciarTablero();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            eliminarCargo();
            cargarCargos();
            reiniciarTablero();
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
            if (!txtNombre.Text.Equals(""))
            {
                if (!txtRiesgo.Text.Equals(""))
                {
                    if (!txtSalario.Text.Equals(""))
                    {
                        modificarCargo();
                        subirArchivo(txtNombre.Text);
                        cargarCargos();
                        reiniciarTablero();
                    }
                    else
                        MessageBox.Show("Favor ingresar el salario del cargo.", "Error");
                }
                else
                    MessageBox.Show("Favor ingresar el riesgo del cargo.", "Error");
            }
            else
                MessageBox.Show("Favor ingresar nombre del cargo.", "Error");
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void buscarCargoLaboral()
        {
            while (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            string query = "SELECT * FROM CargoLaboral ";
            int i = 0;
            if (!txtNombre.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Cargo LIKE '%" + txtNombre.Text + "%'";
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
            if (!txtSalario.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Salario LIKE '" + txtSalario.Text + "'";
            }
            if (!txtFunciones.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Funciones LIKE '%" + txtFunciones.Text + "%'";
            }
            if (!txtRiesgo.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Riesgo LIKE '%" + txtRiesgo.Text + "%'";
            }
            if (checkedListBox1.GetItemCheckState(0).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Casco LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(1).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Tapa_Oidos LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(2).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Careta LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(3).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Lentes LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(4).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Tapa_Boca LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(5).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Delantales LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(6).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Guantes LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(7).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Pantalon_Anticorte LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(8).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Bota_Anticorte LIKE 'Si'";
            }
            if (checkedListBox1.GetItemCheckState(9).ToString().Equals("Checked"))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Botas_Puntera LIKE 'Si'";
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int j = 0;
            try
            {
                while (myReader.Read())
                {
                    double salario = myReader.GetInt32(3);
                    double riesgo = double.Parse(myReader.GetString(5).Replace(".", string.Empty).Replace("%", string.Empty));
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[j].Cells[0].Value = myReader.GetInt32(0);
                    dataGridView1.Rows[j].Cells[1].Value = myReader.GetString(1);
                    dataGridView1.Rows[j].Cells[2].Value = myReader.GetString(2);
                    dataGridView1.Rows[j].Cells[3].Value = String.Format("{0:c}", (salario));
                    dataGridView1.Rows[j].Cells[4].Value = String.Format("{0:c}", (salario * 0.92));
                    dataGridView1.Rows[j].Cells[5].Value = myReader.GetString(4);
                    dataGridView1.Rows[j].Cells[6].Value = myReader.GetString(5) + "%";
                    dataGridView1.Rows[j].Cells[7].Value = String.Format("{0:c}", (salario / 2));
                    dataGridView1.Rows[j].Cells[8].Value = String.Format("{0:c}", (salario / 2));
                    dataGridView1.Rows[j].Cells[9].Value = String.Format("{0:c}", (salario));
                    dataGridView1.Rows[j].Cells[10].Value = String.Format("{0:c}", (salario * 0.12));
                    dataGridView1.Rows[j].Cells[11].Value = String.Format("{0:c}", (salario / 2));
                    double ILA = (salario * 12) + ((salario / 2) * 3) + (salario) + (salario * 0.12);
                    dataGridView1.Rows[j].Cells[12].Value = String.Format("{0:c}", (ILA));
                    dataGridView1.Rows[j].Cells[13].Value = String.Format("{0:c}", (salario * 0.08));
                    dataGridView1.Rows[j].Cells[14].Value = String.Format("{0:c}", (salario * 0.08));
                    dataGridView1.Rows[j].Cells[15].Value = String.Format("{0:c}", (salario * (double.Parse("0.0" + riesgo))));
                    dataGridView1.Rows[j].Cells[16].Value = String.Format("{0:c}", (salario * 0.09));
                    double totalss = (salario * 0.08) + (salario * 0.08) + (salario * double.Parse("0.0" + riesgo)) + (salario * 0.09);
                    double total = (totalss * 12) + ILA;
                    dataGridView1.Rows[j].Cells[17].Value = String.Format("{0:c}", totalss);
                    dataGridView1.Rows[j].Cells[18].Value = String.Format("{0:c}", total);
                    dataGridView1.Rows[j].Cells[19].Value = myReader.GetString(6);
                    dataGridView1.Rows[j].Cells[20].Value = myReader.GetString(7);
                    dataGridView1.Rows[j].Cells[21].Value = myReader.GetString(8);
                    dataGridView1.Rows[j].Cells[22].Value = myReader.GetString(9);
                    dataGridView1.Rows[j].Cells[23].Value = myReader.GetString(10);
                    dataGridView1.Rows[j].Cells[24].Value = myReader.GetString(11);
                    dataGridView1.Rows[j].Cells[25].Value = myReader.GetString(12);
                    dataGridView1.Rows[j].Cells[26].Value = myReader.GetString(13);
                    dataGridView1.Rows[j].Cells[27].Value = myReader.GetString(14);
                    dataGridView1.Rows[j].Cells[28].Value = myReader.GetString(15);
                    j++;
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

        private void button1_Click(object sender, EventArgs e)
        {
            buscarCargoLaboral();
        }

        public void subirArchivo(string cargo) {
            string Descripcion = txtDesc.Text;
            if (!Descripcion.Equals(""))
            {
                Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\" + this.Text + "\\" + cargo);
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\" + this.Text + "\\" + cargo, "Descripcion*");
                if (prueba.Length > 0)
                {
                    if (File.Exists(prueba[0]))
                    {

                        File.Delete(prueba[0]);
                    }
                }
                using (FileStream fs = File.Open(Descripcion, FileMode.Open))
                {
                    Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\" + this.Text + "\\" + cargo);

                    string ext = Path.GetExtension(Descripcion);
                    fs.CopyTo(File.Create("C:\\Users\\" + Environment.UserName + "\\Dropbox\\" + this.Text + "\\" + cargo + "\\Descripcion" + ext));
                }
            }
        }

        private void btnExaminar_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            txtDesc.Text = openFileDialog1.FileName;
        }

        public void modificarSalario()
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE CargoLaboral SET Salario=@Salario WHERE Salario = " + textBox1.Text);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {                
                cmd.Parameters.Add("@Salario", OleDbType.VarChar).Value = textBox2.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Salarios modificado.");
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

        public void imprimirCargos(DataGridView data)
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

        private void button2_Click(object sender, EventArgs e)
        {
            modificarSalario();
            cargarCargos();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            imprimirCargos(dataGridView1);
        }
    }
}
