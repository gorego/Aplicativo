﻿using System;
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
    public partial class frmReciboMaderCliente : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string lote, area, ano, registro, propietario, proveedor, transportador, placa, cliente, FSC;
        int ordenT = 0;

        public frmReciboMaderCliente()
        {
            InitializeComponent();
            cargarLotes();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm"; // Only use hours and minutes
            dateTimePicker2.ShowUpDown = true;
            Variables.cargar(comboBox3, "SELECT * FROM Proveedores", "Proveedor");
            Variables.cargar(comboBox4, "SELECT * FROM Clientes", "Cliente");
            Variables.cargar(comboBox6, "SELECT ID,(Nombres + ' ' + Apellidos) As Nombre FROM Transportadores", "Nombre");
            Variables.cargar(comboBox2, "SELECT * FROM Propietarios", "Nombre");
            Variables.cargar(comboBox5, "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias", "Maquina");
            comboBox1.SelectedItem = null;
            comboBox2.SelectedItem = null;
            comboBox3.SelectedItem = null;
            comboBox4.SelectedItem = null;
            comboBox5.SelectedItem = null;
            comboBox6.SelectedItem = null;
            Variables.cargar(dataGridView2, "SELECT * FROM reciboCliente WHERE volumenActual > 0 ORDER BY ID Desc");
            Variables.cargar(dataGridView1, "SELECT * FROM reciboCliente WHERE (Month(Fecha) BETWEEN " + (DateTime.Now.Month - 3) + " AND " + (DateTime.Now.Month) + ") ORDER BY ID Desc");
            Variables.cargar(dataGridView3, "SELECT * FROM reciboCliente WHERE Especie = 'Melina' AND volumenActual > 0 ORDER BY ID Desc");
            Variables.cargar(dataGridView4, "SELECT * FROM reciboCliente WHERE Especie = 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
            Variables.cargar(dataGridView5, "SELECT * FROM reciboCliente WHERE Especie <> 'Melina' AND Especie <> 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
            crearFormatoData(dataGridView5);
            crearFormatoData(dataGridView4);
            crearFormatoData(dataGridView3);
            crearFormatoData(dataGridView2);
            crearFormatoData(dataGridView1);
            txtModulo.Text = getMaxPila().ToString();
            label16.Text = "Volumen Total:  " + getTotal(dataGridView2) + " m3.";
        }

        public int getMaxPila()
        {
            string query = "SELECT MAX(Pila) From reciboCliente";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            int pila = 0;
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {                    
                    if(!myReader.GetValue(0).ToString().Equals(""))
                        pila = Int32.Parse(myReader.GetValue(0).ToString());
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return pila;
        }

        public int getOrdenID()
        {
            string orden = "OT-" + txtOT.Text.Trim().PadLeft(4, '0') + "-" + textBox7.Text.Trim();                    
            string query = "SELECT ID From historicoOrdenes WHERE OT = '" + orden + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            int pila = 0;
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    if (!myReader.GetValue(0).ToString().Equals(""))
                        pila = Int32.Parse(myReader.GetValue(0).ToString());
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return pila;
        }

        public void cargarLotes()
        {
            string query = "SELECT Codigo, Lote FROM Lotes Group By Codigo,Lote UNION ALL SELECT Codigo, Lote FROM Areas Group By Codigo,Lote UNION ALL SELECT Codigo, Lote FROM LoteGanadero Group By Codigo,Lote";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.DisplayMember = "Lote";
            comboBox1.ValueMember = "Codigo";
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void crearFormatoData(DataGridView dataGridView2)
        {
            dataGridView2.Columns[0].HeaderText = "# Recibo";
            dataGridView2.Columns[1].Visible = false;
            dataGridView2.Columns[2].Visible = false;
            dataGridView2.Columns[3].Visible = false;
            //dataGridView2.Columns[4].Visible = false;
            dataGridView2.Columns[5].HeaderText = "Volumen Ingresado";
            dataGridView2.Columns[6].HeaderText = "Volumen Actual";
            //dataGridView2.Columns[7].Visible = false;
            //dataGridView2.Columns[8].Visible = false;
            dataGridView2.Columns[9].Visible = false;
            dataGridView2.Columns[10].Visible = false;
            dataGridView2.Columns[11].Visible = false;
            dataGridView2.Columns[12].Visible = false;
            dataGridView2.Columns[13].Visible = false;
            dataGridView2.Columns[14].Visible = false;
            dataGridView2.Columns[15].Visible = false;
            dataGridView2.Columns[16].Visible = false;
            dataGridView2.Columns[17].Visible = false;
            dataGridView2.Columns[18].Visible = false;
            dataGridView2.Columns[19].Visible = false;
            dataGridView2.Columns[20].Visible = false;
            dataGridView2.Columns[21].Visible = false;
            dataGridView2.Columns[22].Visible = false;
            dataGridView2.Columns[25].Visible = false;
            dataGridView2.Columns[24].HeaderText = "# Recibo";
            dataGridView2.Columns[24].DisplayIndex = 0;
        }

        private void txtOT_Leave(object sender, EventArgs e)
        {
            if (textBox7.Text.Equals(""))
            {
                //MessageBox.Show("Favor seleccionar el año de la OT.");
            }
            else
            {
                int prueba = 0;
                bool isNum = Int32.TryParse(txtOT.Text.Trim(), out prueba);
                if (isNum)
                {
                    string orden = "OT-" + txtOT.Text.Trim().PadLeft(4, '0') + "-" + textBox7.Text.Trim();                    
                    if (Variables.existe("SELECT * FROM historicoOrdenes WHERE OT = '" + orden + "'"))
                    {
                        getOrden((orden));
                        comboBox1.SelectedValue = lote;
                        textBox2.Text = area;
                        if (ano.Equals(""))
                            textBox3.Text = "0";
                        else
                            textBox3.Text = ano;
                        comboBox6.SelectedValue = transportador;
                        comboBox4.SelectedValue = cliente;
                        comboBox2.SelectedValue = propietario;
                        comboBox3.SelectedValue = proveedor;
                        if (registro.Equals(""))
                            textBox4.Text = "0";
                        else
                            textBox4.Text = registro;
                        textBox1.Text = placa;
                        if (getEquipo(ordenT) != 0)
                            comboBox5.SelectedValue = getEquipo(ordenT);
                        if (FSC.Equals("0"))
                            radioButton14.Checked = true;
                        else
                        {
                            radioButton13.Checked = true;
                            textBox6.Text = FSC;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Orden de Trabajo no existe.", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("Favor digitar un numero valido.", "Error");
                }
            }
        }

        public int getEquipo(int orden)
        {
            int equipo = 0;
            string query = "SELECT Maquina FROM ordenMaquinas WHERE Orden = " + orden;
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
                    equipo = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return equipo;
        }

        public void getOrden(string orden)
        {
            string query = "SELECT h.Lote, h.Area, l.Ano, l.registroPlantacion, b.Propietario, t.Proveedor, h.Transportador, t.Placa, h.Cliente, l.FSC,h.ID FROM BancoTierras AS b INNER JOIN ((Transportadores AS t INNER JOIN historicoOrdenes AS h ON t.ID = h.Transportador) INNER JOIN Lotes AS l ON h.Lote = l.Codigo) ON b.ID = l.Predio WHERE h.OT = '" + orden + "'";
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
                    lote = myReader.GetValue(0).ToString();
                    area = myReader.GetValue(1).ToString();
                    ano = myReader.GetValue(2).ToString();
                    registro = myReader.GetValue(3).ToString();
                    propietario = myReader.GetValue(4).ToString();
                    proveedor = myReader.GetValue(5).ToString();
                    transportador = myReader.GetValue(6).ToString();
                    placa = myReader.GetValue(7).ToString();
                    cliente = myReader.GetValue(8).ToString();
                    FSC = myReader.GetValue(9).ToString();
                    ordenT = myReader.GetInt32(10);
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

        private void radioButton28_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton28.Checked)
                textBox36.Visible = true;
            else
                textBox36.Visible = false;
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked)
            {
                textBox6.Visible = true;
                label76.Visible = true;
            }
            else
            {
                textBox6.Visible = false;
                label76.Visible = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            txtModulo.Text = (Int32.Parse(txtModulo.Text) + 1) + "";
        }

        public int getMaxID()
        {
            string query = "SELECT MAX(numPila) FROM reciboCliente WHERE pila = " + txtModulo.Text;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int id = 0;
            string value = "";
            try
            {
                if (myReader.Read())
                {
                    value = myReader.GetValue(0).ToString();
                    if (value.Equals(""))
                        id = 0;
                    else
                        id = Int32.Parse(value);
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

        public bool existeAñoActual()
        {
            string query = "SELECT * FROM reciboCliente WHERE numRecibo like '%" + textBox7.Text + "%'";
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
                    return true;
                }
                else
                {
                    return false;
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

        public void agregarRecibo(int id,int ot)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("INSERT INTO reciboCliente (Orden,Clasificacion,Motivo,Especie,volumenIngreso,volumenActual,Diametro,Largo,Cantidad,Lote,Ano,Area,Registro,Propietario,Proveedor,Conductor,Placa,Cliente,FSC,Fecha,Hora,Maquinaria,Pila,numRecibo,numPila) VALUES (@Orden,@Clasificacion,@Motivo,@Especie,@volumenIngreso,@volumenActual,@Diametro,@Largo,@Cantidad,@Lote,@Ano,@Area,@Registro,@Propietario,@Proveedor,@Conductor,@Placa,@Cliente,@FSC,@Fecha,@Hora,@Maquinaria,@Pila,@numRecibo,@numPila)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = ot;
                if (radioButton1.Checked)
                    cmd.Parameters.Add("@Clasificacion", OleDbType.VarChar).Value = "Transferencia de Material";
                else if (radioButton2.Checked)
                    cmd.Parameters.Add("@Clasificacion", OleDbType.VarChar).Value = "Compra de Material";
                else
                    cmd.Parameters.Add("@Clasificacion", OleDbType.VarChar).Value = "Prestacion de Servicio";
                if (radioButton6.Checked)
                    cmd.Parameters.Add("@Motivo", OleDbType.VarChar).Value = "Entresaca";
                else if (radioButton5.Checked)
                    cmd.Parameters.Add("@Motivo", OleDbType.VarChar).Value = "Tala Raza";
                else
                    cmd.Parameters.Add("@Motivo", OleDbType.VarChar).Value = "Recuperacion de Material";
                if (radioButton18.Checked)
                    cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Melina";
                else if (radioButton17.Checked)
                    cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Teca";
                else
                    cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = textBox36.Text;
                cmd.Parameters.Add("@volumenIngreso", OleDbType.VarChar).Value = textBox38.Text.Replace(".", ",");
                cmd.Parameters.Add("@volumenActual", OleDbType.VarChar).Value = textBox38.Text.Replace(".", ",");
                cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox41.Text;
                cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = textBox40.Text;
                cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox39.Text;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Registro", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                cmd.Parameters.Add("@Conductor", OleDbType.VarChar).Value = comboBox6.SelectedValue;
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = comboBox4.SelectedValue;
                cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = dateTimePicker2.Value.Hour + ":" + dateTimePicker2.Value.Minute;
                cmd.Parameters.Add("@Maquinaria", OleDbType.VarChar).Value = comboBox5.SelectedValue;
                cmd.Parameters.Add("@Pila", OleDbType.VarChar).Value = txtModulo.Text;
                cmd.Parameters.Add("@numRecibo", OleDbType.VarChar).Value = "PILA-" + txtModulo.Text + "-" + id.ToString().PadLeft(4, '0') + "-OT-" + txtOT.Text.Trim().PadLeft(4,'0')+ "-" + textBox7.Text;
                cmd.Parameters.Add("@numPila", OleDbType.VarChar).Value = id;
                try
                {
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("Materia Prima Recibida");
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals(""))
            {
                MessageBox.Show("Favor seleccionar lote.");
            }
            else
            {
                if (textBox3.Text.Equals(""))
                {
                    MessageBox.Show("Favor seleccionar año de plantación.");
                }
                else
                {
                    if (textBox2.Text.Equals(""))
                    {
                        MessageBox.Show("Favor seleccionar area.");
                    }
                    else
                    {
                        if (comboBox2.Text.Equals(""))
                        {
                            MessageBox.Show("Favor seleccionar propietario.");
                        }
                        else
                        {
                            if (comboBox3.Text.Equals(""))
                            {
                                MessageBox.Show("Favor seleccionar proveedor.");
                            }
                            else
                            {
                                if (comboBox6.Text.Equals(""))
                                {
                                    MessageBox.Show("Favor seleccionar conductor.");
                                }
                                else
                                {
                                    if (comboBox4.Text.Equals(""))
                                    {
                                        MessageBox.Show("Favor seleccionar cliente.");
                                    }
                                    else
                                    {
                                        if (comboBox5.Text.Equals(""))
                                        {
                                            MessageBox.Show("Favor seleccionar maquinaria.");
                                        }
                                        else
                                        {
                                            if (txtOT.Text.Equals(""))
                                            {
                                                MessageBox.Show("Favor seleccionar OT.");
                                            }
                                            else
                                            {
                                                int id = getMaxID();
                                                if (!existeAñoActual())
                                                    id = 0;
                                                agregarRecibo(id + 1, getOrdenID());
                                                Variables.cargar(dataGridView2, "SELECT * FROM reciboCliente WHERE volumenActual > 0 ORDER BY ID Desc");
                                                Variables.cargar(dataGridView1, "SELECT * FROM reciboCliente WHERE (Month(Fecha) BETWEEN " + (DateTime.Now.Month - 3) + " AND " + (DateTime.Now.Month) + ") ORDER BY ID Desc");
                                                Variables.cargar(dataGridView3, "SELECT * FROM reciboCliente WHERE Especie = 'Melina' AND volumenActual > 0 ORDER BY ID Desc");
                                                Variables.cargar(dataGridView4, "SELECT * FROM reciboCliente WHERE Especie = 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
                                                Variables.cargar(dataGridView5, "SELECT * FROM reciboCliente WHERE Especie <> 'Melina' AND Especie <> 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtOT.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString();
            if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString().Equals("Transferencia de Material"))
                radioButton1.Checked = true;
            else if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[2].Value.ToString().Equals("Compra de Material"))
                radioButton2.Checked = true;
            else
                radioButton3.Checked = true;
            if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString().Equals("Entresaca"))
                radioButton6.Checked = true;
            else if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[3].Value.ToString().Equals("Tala Raza"))
                radioButton5.Checked = true;
            else
                radioButton7.Checked = true;
            if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[4].Value.ToString().Equals("Melina"))
                radioButton18.Checked = true;
            else if (dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[4].Value.ToString().Equals("Teca"))
                radioButton17.Checked = true;
            else
            {
                radioButton28.Checked = true;
                textBox36.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[4].Value.ToString();
            }
            textBox38.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox5.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[6].Value.ToString();
            textBox41.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox40.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox39.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[9].Value.ToString();
            comboBox1.SelectedValue = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox3.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[11].Value.ToString();
            textBox2.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox4.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[13].Value.ToString();
            comboBox2.SelectedValue = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[14].Value.ToString();
            comboBox3.SelectedValue = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[15].Value.ToString();
            comboBox6.SelectedValue = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[16].Value.ToString();
            textBox1.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[17].Value.ToString();
            comboBox4.SelectedValue = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[18].Value.ToString();
            textBox6.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[19].Value.ToString();
            if (textBox6.Equals("0"))
                radioButton14.Checked = true;
            dateTimePicker1.Value = DateTime.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[20].Value.ToString());
            dateTimePicker2.Value = DateTime.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[21].Value.ToString());
            comboBox5.SelectedValue = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[22].Value.ToString();
            txtModulo.Text = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[23].Value.ToString();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtOT.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString().Equals("Transferencia de Material"))
                radioButton1.Checked = true;
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString().Equals("Compra de Material"))
                radioButton2.Checked = true;
            else
                radioButton3.Checked = true;
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString().Equals("Entresaca"))
                radioButton6.Checked = true;
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString().Equals("Tala Raza"))
                radioButton5.Checked = true;
            else
                radioButton7.Checked = true;
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString().Equals("Melina"))
                radioButton18.Checked = true;
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString().Equals("Teca"))
                radioButton17.Checked = true;
            else
            {
                radioButton28.Checked = true;
                textBox36.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            }
            textBox38.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            textBox41.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox40.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox39.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            comboBox1.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            comboBox2.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            comboBox3.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            comboBox6.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            comboBox4.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            textBox6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[19].Value.ToString();
            if (textBox6.Equals("0"))
                radioButton14.Checked = true;
            dateTimePicker1.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[20].Value.ToString());
            dateTimePicker2.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[21].Value.ToString());
            comboBox5.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[22].Value.ToString();
            txtModulo.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[23].Value.ToString();
        }

        public void clickDataGridView(DataGridView dataGridView1)
        {
            txtOT.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString();
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString().Equals("Transferencia de Material"))
                radioButton1.Checked = true;
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString().Equals("Compra de Material"))
                radioButton2.Checked = true;
            else
                radioButton3.Checked = true;
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString().Equals("Entresaca"))
                radioButton6.Checked = true;
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[3].Value.ToString().Equals("Tala Raza"))
                radioButton5.Checked = true;
            else
                radioButton7.Checked = true;
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString().Equals("Melina"))
                radioButton18.Checked = true;
            else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString().Equals("Teca"))
                radioButton17.Checked = true;
            else
            {
                radioButton28.Checked = true;
                textBox36.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[4].Value.ToString();
            }
            textBox38.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[5].Value.ToString();
            textBox5.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString();
            textBox41.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString();
            textBox40.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
            textBox39.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
            comboBox1.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
            textBox3.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[11].Value.ToString();
            textBox2.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[12].Value.ToString();
            textBox4.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[13].Value.ToString();
            comboBox2.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[14].Value.ToString();
            comboBox3.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[15].Value.ToString();
            comboBox6.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[16].Value.ToString();
            textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[17].Value.ToString();
            comboBox4.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[18].Value.ToString();
            textBox6.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[19].Value.ToString();
            if (textBox6.Equals("0"))
                radioButton14.Checked = true;
            dateTimePicker1.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[20].Value.ToString());
            dateTimePicker2.Value = DateTime.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[21].Value.ToString());
            comboBox5.SelectedValue = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[22].Value.ToString();
            txtModulo.Text = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[23].Value.ToString();
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            clickDataGridView(dataGridView3);
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            clickDataGridView(dataGridView4);
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            clickDataGridView(dataGridView5);
        }

        public void modificarRecibo(string recibo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE reciboCliente SET Orden=@Orden,Clasificacion=@Clasificacion,Motivo=@Motivo,Especie=@Especie,volumenIngreso=@volumenIngreso,volumenActual=@volumenActual,Diametro=@Diametro,Largo=@Largo,Cantidad=@Cantidad,Lote=@Lote,Ano=@Ano,Area=@Area,Registro=@Registro,Propietario=@Propietario,Proveedor=@Proveedor,Conductor=@Conductor,Placa=@Placa,Cliente=@Cliente,FSC=@FSC,Fecha=@Fecha,Hora=@Hora,Maquinaria=@Maquinaria,Pila=@Pila WHERE ID = " + recibo);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = txtOT.Text;
                if (radioButton1.Checked)
                    cmd.Parameters.Add("@Clasificacion", OleDbType.VarChar).Value = "Transferencia de Material";
                else if (radioButton2.Checked)
                    cmd.Parameters.Add("@Clasificacion", OleDbType.VarChar).Value = "Compra de Material";
                else
                    cmd.Parameters.Add("@Clasificacion", OleDbType.VarChar).Value = "Prestacion de Servicio";
                if (radioButton6.Checked)
                    cmd.Parameters.Add("@Motivo", OleDbType.VarChar).Value = "Entresaca";
                else if (radioButton5.Checked)
                    cmd.Parameters.Add("@Motivo", OleDbType.VarChar).Value = "Tala Raza";
                else
                    cmd.Parameters.Add("@Motivo", OleDbType.VarChar).Value = "Recuperacion de Material";
                if (radioButton18.Checked)
                    cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Melina";
                else if (radioButton17.Checked)
                    cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Teca";
                else
                    cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = textBox36.Text;
                cmd.Parameters.Add("@volumenIngreso", OleDbType.VarChar).Value = textBox38.Text.Replace(".", ",");
                cmd.Parameters.Add("@volumenActual", OleDbType.VarChar).Value = textBox5.Text.Replace(".", ",");
                cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox41.Text;
                cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = textBox40.Text;
                cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox39.Text;
                cmd.Parameters.Add("@Lote", OleDbType.VarChar).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@Ano", OleDbType.VarChar).Value = textBox3.Text;
                cmd.Parameters.Add("@Area", OleDbType.VarChar).Value = textBox2.Text;
                cmd.Parameters.Add("@Registro", OleDbType.VarChar).Value = textBox4.Text;
                cmd.Parameters.Add("@Propietario", OleDbType.VarChar).Value = comboBox2.SelectedValue;
                cmd.Parameters.Add("@Proveedor", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                cmd.Parameters.Add("@Conductor", OleDbType.VarChar).Value = comboBox6.SelectedValue;
                cmd.Parameters.Add("@Placa", OleDbType.VarChar).Value = textBox1.Text;
                cmd.Parameters.Add("@Cliente", OleDbType.VarChar).Value = comboBox4.SelectedValue;
                cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = dateTimePicker1.Value.ToString("dd") + "/" + dateTimePicker1.Value.ToString("MM") + "/" + dateTimePicker1.Value.Year;
                cmd.Parameters.Add("@Hora", OleDbType.VarChar).Value = dateTimePicker2.Value.Hour + ":" + dateTimePicker2.Value.Minute;
                cmd.Parameters.Add("@Maquinaria", OleDbType.VarChar).Value = comboBox5.SelectedValue;
                cmd.Parameters.Add("@Modulo", OleDbType.VarChar).Value = txtModulo.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("Recibo Modificado.");
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

        private void button3_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
                modificarRecibo(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());
            else if (tabControl1.SelectedIndex == 1)
                modificarRecibo(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
            else if (tabControl1.SelectedIndex == 2)
                modificarRecibo(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString());
            else if (tabControl1.SelectedIndex == 3)
                modificarRecibo(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString());
            else
                modificarRecibo(dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[0].Value.ToString());
            Variables.cargar(dataGridView2, "SELECT * FROM reciboCliente WHERE volumenActual > 0 ORDER BY ID Desc");
            Variables.cargar(dataGridView1, "SELECT * FROM reciboCliente WHERE (Month(Fecha) BETWEEN " + (DateTime.Now.Month - 3) + " AND " + (DateTime.Now.Month) + ") ORDER BY ID Desc");
            Variables.cargar(dataGridView3, "SELECT * FROM reciboCliente WHERE Especie = 'Melina' AND volumenActual > 0 ORDER BY ID Desc");
            Variables.cargar(dataGridView4, "SELECT * FROM reciboCliente WHERE Especie = 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
            Variables.cargar(dataGridView5, "SELECT * FROM reciboCliente WHERE Especie <> 'Melina' AND Especie <> 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = new DialogResult();
            if (tabControl1.SelectedIndex == 0)
                dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el recibo # " + dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            else if (tabControl1.SelectedIndex == 1)
                dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el recibo # " + dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            else if (tabControl1.SelectedIndex == 2)
                dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el recibo # " + dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            else if (tabControl1.SelectedIndex == 3)
                dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el recibo # " + dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            else
                dialogResult = System.Windows.Forms.MessageBox.Show("Seguro de eliminar el recibo # " + dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[1].Value.ToString() + "?", "Confirmar", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string id = "";
                if (tabControl1.SelectedIndex == 0)
                    id = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
                else
                    id = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("DELETE FROM reciboCliente WHERE id = " + id);
                cmd.Connection = conn;
                conn.Open();

                if (conn.State == ConnectionState.Open)
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Recibo eliminado.");
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
                Variables.cargar(dataGridView2, "SELECT * FROM reciboCliente WHERE volumenActual > 0 ORDER BY ID Desc");
                Variables.cargar(dataGridView1, "SELECT * FROM reciboCliente WHERE (Month(Fecha) BETWEEN " + (DateTime.Now.Month - 3) + " AND " + (DateTime.Now.Month) + ") ORDER BY ID Desc");
                Variables.cargar(dataGridView3, "SELECT * FROM reciboCliente WHERE Especie = 'Melina' AND volumenActual > 0 ORDER BY ID Desc");
                Variables.cargar(dataGridView4, "SELECT * FROM reciboCliente WHERE Especie = 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
                Variables.cargar(dataGridView5, "SELECT * FROM reciboCliente WHERE Especie <> 'Melina' AND Especie <> 'Teca' AND volumenActual > 0 ORDER BY ID Desc");
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            if (txtOT.Text.Equals(""))
            {
                //MessageBox.Show("Favor seleccionar el # de la OT.");
            }
            else
            {
                int prueba = 0;
                bool isNum = Int32.TryParse(textBox7.Text.Trim(), out prueba);
                if (isNum)
                {
                    string orden = "OT-" + txtOT.Text.Trim().PadLeft(4, '0') + "-" + textBox7.Text.Trim();
                    if (Variables.existe("SELECT * FROM historicoOrdenes WHERE OT = '" + orden + "'"))
                    {
                        getOrden(orden);
                        comboBox1.SelectedValue = lote;
                        textBox2.Text = area;
                        if (ano.Equals(""))
                            textBox3.Text = "0";
                        else
                            textBox3.Text = ano;
                        comboBox6.SelectedValue = transportador;
                        comboBox4.SelectedValue = cliente;
                        comboBox2.SelectedValue = propietario;
                        comboBox3.SelectedValue = proveedor;
                        if (registro.Equals(""))
                            textBox4.Text = "0";
                        else
                            textBox4.Text = registro;
                        textBox1.Text = placa;
                        if (getEquipo(ordenT) != 0)
                            comboBox5.SelectedValue = getEquipo(ordenT);
                        if (FSC.Equals("0"))
                            radioButton14.Checked = true;
                        else
                        {
                            radioButton13.Checked = true;
                            textBox6.Text = FSC;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Orden de Trabajo no existe.", "Error");
                    }
                }
                else
                {
                    MessageBox.Show("Favor digitar un numero valido.", "Error");
                }
            }
        }

        public string getTotal(DataGridView data)
        {
            string total = "";
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[6].Value.ToString());
            }
            total = valor.ToString();
            return total;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                label16.Text = "Volumen Total:  " + getTotal(dataGridView2) + " m3."; 
            }
            else if (tabControl1.SelectedIndex == 1) 
            {
                label16.Text = "Volumen Total:  " + getTotal(dataGridView1) + " m3."; 
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                label16.Text = "Volumen Total:  " + getTotal(dataGridView3) + " m3.";
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                label16.Text = "Volumen Total:  " + getTotal(dataGridView4) + " m3.";
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                label16.Text = "Volumen Total:  " + getTotal(dataGridView5) + " m3.";
            }
        }

        public double getVolumen(double diametro, double largo, int cantidad, string especie, string trailer, string raleo, string adf)
        {
            double volumen = 0;
            int count = 0;
            string query = "SELECT Volumen FROM volumenCalculado WHERE Diametro = " + diametro.ToString().Replace(",", ".") + " AND Largo = " + largo.ToString().Replace(",", ".") + " AND Especie = '" + especie + "' AND Trailer = '" + trailer + "' and Raleo = '" + raleo + "'";
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
                    volumen += double.Parse(myReader.GetValue(0).ToString()) * cantidad;
                    count++;
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            double vol = 0;
            vol = Math.Round(volumen / count, 4, MidpointRounding.AwayFromZero);
            if (volumen / count == 0)
                return 0;
            else
                return vol;
        }

        public string getRaleo(RadioButton r1, RadioButton r2, RadioButton r3)
        {
            string raleo = "";
            if (r1.Checked)
                raleo = "Entresaca";
            else if (r2.Checked)
                raleo = "Tala Raza";
            else
                raleo = "Recuperacion de Material";
            return raleo;
        }

        public string getTrailer(RadioButton r1, RadioButton r2, RadioButton r3, RadioButton r4, RadioButton r5)
        {
            string trailer = "";
            if (r1.Checked)
                trailer = "Farmi Primero 9000";
            else if (r2.Checked)
                trailer = "Pfanzelt 15100";
            else if (r3.Checked)
                trailer = "Camion 600";
            else if (r4.Checked)
                trailer = "Camion Doble Torque";
            else
                trailer = "Tractomula";
            return trailer;
        }

        public string getEspecie(RadioButton r1, RadioButton r2, RadioButton r3, TextBox otro)
        {
            string especie = "";
            if (r1.Checked)
                especie = "Melina";
            else if (r2.Checked)
                especie = "Teca";
            else
                especie = otro.Text;
            return especie;
        }

        private void textBox38_Enter(object sender, EventArgs e)
        {
            if (!textBox41.Text.Equals("") && !textBox40.Text.Equals("") && !textBox39.Text.Equals(""))
            {
                string raleo = "", especie = "", trailer = "";
                raleo = getRaleo(radioButton6, radioButton5, radioButton7);
                especie = getEspecie(radioButton18, radioButton17, radioButton28, textBox36);
                trailer = getTrailer(radioButton9, radioButton8, radioButton4, radioButton11, radioButton10);
                textBox38.Text = getVolumen(double.Parse(textBox41.Text), double.Parse(textBox40.Text), Int32.Parse(textBox39.Text), especie, trailer, raleo, "ADF012").ToString();
            }
        }

    }
}
