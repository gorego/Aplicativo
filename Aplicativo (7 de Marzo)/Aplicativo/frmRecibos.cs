﻿using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace Aplicativo
{
    public partial class frmRecibos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string mod = "";
        int view = 0;

        public frmRecibos()
        {
            InitializeComponent();
            Variables.cargar(dataGridView2, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE (((Recibo.volumenActual)>0)) ORDER BY Recibo.ID DESC;");
            Variables.cargar(dataGridView1, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID ORDER BY Recibo.ID DESC;");
            Variables.cargar(dataGridView3, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE Recibo.Especie = 'Melina' ORDER BY Recibo.ID DESC;");
            Variables.cargar(dataGridView4, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE Recibo.Especie = 'Teca' ORDER BY Recibo.ID DESC;");
            Variables.cargar(dataGridView5, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE Recibo.Especie <> 'Melina' AND Recibo.Especie <> 'Teca' ORDER BY Recibo.ID DESC;");
            crearFormatoData(dataGridView1);
            crearFormatoData(dataGridView2);
            crearFormatoData(dataGridView3);
            crearFormatoData(dataGridView4);
            crearFormatoData(dataGridView5);
            cargarLotes(comboBox5);
            comboBox5.SelectedItem = null;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd/MM/yyyy";
        }

        public void cargarLotes(ComboBox txtLote)
        {
            string query = "SELECT Codigo, Lote FROM Lotes";
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            txtLote.DataSource = ds.Tables[0];
            txtLote.DisplayMember = "Lote";
            txtLote.ValueMember = "Codigo";
            txtLote.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtLote.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void crearFormatoData(DataGridView dataGridView2)
        {
            dataGridView2.Columns[0].HeaderText = "# Recibo";
            dataGridView2.Columns[1].Visible = false;
            //dataGridView2.Columns[4].Visible = false;
            dataGridView2.Columns[5].HeaderText = "Volumen Ingresado";
            dataGridView2.Columns[6].HeaderText = "Volumen Actual";
            //dataGridView2.Columns[7].Visible = false;
            //dataGridView2.Columns[8].Visible = false;            
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
            dataGridView2.Columns[22].Visible = false;
            dataGridView2.Columns[24].HeaderText = "# Recibo";
            dataGridView2.Columns[24].DisplayIndex = 0;
            dataGridView2.Columns[25].HeaderText = "Lote";
            dataGridView2.Columns[dataGridView2.Columns.Count - 1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView2.Columns[24].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            buscarRecibo("(((Recibo.volumenActual)>0))", dataGridView2);
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            abrirOrden(dataGridView2);
            view = 2;
        }

        public void abrirOrden(DataGridView data) 
        {
            if (data.CurrentCell.ColumnIndex == 26)
            {
                frmCrearOrden newFrm = new frmCrearOrden(data.Rows[data.CurrentCell.RowIndex].Cells[1].Value.ToString(), 1);
                newFrm.Show();
            }
            else if (data.CurrentCell.ColumnIndex == 24)
            {
                comboBox4.Text = data.Rows[data.CurrentCell.RowIndex].Cells[23].Value.ToString();
                button2.Text = "Modificar Modulo Recibo " + data.Rows[data.CurrentCell.RowIndex].Cells[24].Value.ToString();
                mod = data.Rows[data.CurrentCell.RowIndex].Cells[0].Value.ToString();                
                //reciboModulo newFrm = new reciboModulo(data.Rows[data.CurrentCell.RowIndex].Cells[0].Value.ToString(), data.Rows[data.CurrentCell.RowIndex].Cells[24].Value.ToString(), data.Rows[data.CurrentCell.RowIndex].Cells[23].Value.ToString());
                //newFrm.Show();
            }
            else
            {
                button2.Text = "Modificar Modulo";
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            abrirOrden(dataGridView1);
            view = 1;
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            abrirOrden(dataGridView3);
            view = 3;
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            abrirOrden(dataGridView4);
            view = 4;
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            abrirOrden(dataGridView5);
            view = 5;
        }

        public void buscarRecibo(string lastQuery, DataGridView data)
        {
            string query = "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID  ";
            int i = 0;
            if (!comboBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Clasificacion LIKE '%" + comboBox1.Text + "%'";
            }
            if (!comboBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Motivo LIKE '%" + comboBox2.Text + "%'";
            }
            if (!comboBox3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Especie LIKE '%" + comboBox3.Text + "%'";
            }
            if (!comboBox4.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Modulo = " + comboBox4.Text;
            }
            if (!textBox1.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.volumenIngreso = " + textBox1.Text;
            }
            if (!textBox2.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.volumenActual = " + textBox2.Text;
            }
            if (!textBox4.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Diametro = " + textBox4.Text;
            }
            if (!textBox3.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Largo = " + textBox5.Text;
            }
            if (!comboBox5.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Lote = " + comboBox5.SelectedValue;
            }
            if (!textBox5.Text.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += "Recibo.Cantidad = " + textBox5.Text;
            }
            if (!lastQuery.Equals(""))
            {
                if (i != 0)
                    query += " AND ";
                else
                    query += "WHERE ";
                i++;
                query += lastQuery;
            }
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable supervisores = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(supervisores);
            data.DataSource = supervisores;
        }

        private void frmRecibos_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text != "Modificar Modulo")
            {
                if(!mod.Equals(""))
                    modificarModulo(mod);

                if (view == 1)
                    Variables.cargar(dataGridView1, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID ORDER BY Recibo.ID DESC;");
                else if(view == 2)
                    Variables.cargar(dataGridView2, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE (((Recibo.volumenActual)>0)) ORDER BY Recibo.ID DESC;");
                else if (view ==3)
                    Variables.cargar(dataGridView3, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE Recibo.Especie = 'Melina' ORDER BY Recibo.ID DESC;");
                else if (view == 4)
                    Variables.cargar(dataGridView4, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE Recibo.Especie = 'Teca' ORDER BY Recibo.ID DESC;");
                else if(view == 5)
                    Variables.cargar(dataGridView5, "SELECT Recibo.*, Lotes.Lote, h.OT FROM (Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote) INNER JOIN historicoOrdenes AS h ON Recibo.Orden = h.ID WHERE Recibo.Especie <> 'Melina' AND Recibo.Especie <> 'Teca' ORDER BY Recibo.ID DESC;");
            }
            else
            {
                MessageBox.Show("Favor seleccionar recibo.");
            }
        }

        public void modificarModulo(string id)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Recibo SET Modulo=@Modulo WHERE ID = " + id);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Modulo", OleDbType.VarChar).Value = comboBox4.Text;
                try
                {
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Modulo modficiado.");
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

        private void button3_Click(object sender, EventArgs e)
        {
            ReiniciarTablero();
        }

        public void ReiniciarTablero()
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            textBox1.Text = "";            
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            button2.Text = "Modificar Modulo";
        }
    }
}