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
    public partial class frmRegistrarPaquete : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        int end = 0;
        int numEntradas = 0;

        public frmRegistrarPaquete(int op)
        {
            InitializeComponent();
            this.Text = "Registrar Paquete OP # " + getNombreOP(op);
            Variables.cargar(comboBox1, "SELECT op.Id, op.Producto, p.Codigo, op.Cantidad, op.Volumen FROM Productos AS p INNER JOIN produccionProducto AS op ON p.ID = op.Producto WHERE Orden = " + op, "Codigo");
            comboBox1.SelectedItem = null;
            end = 1;
        }

        public string getNombreOP(int op)
        {
            string query = "SELECT OP FROM historicoProduccion WHERE id = " + op;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string value = "";
            try
            {
                if (myReader.Read())
                {
                    value = myReader.GetValue(0).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return value;
        }

        public void getProducto(string op)
        {
            string query = "SELECT numAnchoEmp,numAltoEmp FROM Productos WHERE Codigo = '" + op + "'";
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
                    textBox23.Text = myReader.GetValue(0).ToString();
                    textBox22.Text = myReader.GetValue(1).ToString();
                    textBox13.Text = (myReader.GetDouble(0) * myReader.GetDouble(1)).ToString();
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


        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                textBox1.ReadOnly = false;
            else
            {
                textBox1.ReadOnly = true;
                textBox1.Text = "0";
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (end == 1)           
                getProducto(comboBox1.Text);           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!textBox2.Text.Equals("0"))
            {
                if (!textBox3.Text.Equals("0"))
                {
                    if(!textBox4.Text.Equals("0"))
                    {
                        if(!textBox5.Text.Equals("0"))
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[numEntradas].Cells[0].Value = textBox2.Text;
                            dataGridView1.Rows[numEntradas].Cells[1].Value = textBox3.Text;
                            dataGridView1.Rows[numEntradas].Cells[2].Value = textBox4.Text;
                            dataGridView1.Rows[numEntradas].Cells[3].Value = textBox5.Text;
                            numEntradas++;
                        }
                        else
                            MessageBox.Show("Favor insertar un numero superior a 0 en el Alto 2.");
                    }
                    else
                            MessageBox.Show("Favor insertar un numero superior a 0 en el Alto 1.");
                }
                else
                    MessageBox.Show("Favor insertar un numero superior a 0 en el Ancho 2.");
            }
            else
                MessageBox.Show("Favor insertar un numero superior a 0 en el Ancho 1.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (numEntradas < 5)
            {
                MessageBox.Show("Favor llenar 5 registros.");
            }
        }

    }
}
