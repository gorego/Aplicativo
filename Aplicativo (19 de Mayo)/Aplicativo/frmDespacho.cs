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

namespace Aplicativo
{
    public partial class frmDespacho : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmDespacho()
        {
            InitializeComponent();
            cargar(dataGridView2,DateTime.Now.Year);
        }

        public void cargar(DataGridView data, int ano)
        {
            if (Variables.tipo == 2)
            {
                if (ano != -1)
                    Variables.cargar(data, "SELECT Despacho.ID,Despacho.despacho, Clientes.cliente, (Transportadores.nombres + ' ' + Transportadores.apellidos) As Transportador, numICA, numICAWeb, FSC, Fecha, Hora, Tipo FROM (Despacho INNER JOIN Transportadores ON Despacho.Transportador = Transportadores.ID) INNER JOIN Clientes ON Despacho.Cliente = Clientes.ID WHERE (((Year([Fecha]))=" + ano + ")) AND Despacho.Tipo = 'Madera Rolliza' ORDER BY Despacho.ID DESC;");
                else
                    Variables.cargar(data, "SELECT Despacho.ID, Despacho.despacho, Clientes.cliente, Transportadores.nombres, numICA, numICAWeb, FSC, Fecha, Hora, Tipo FROM (Despacho INNER JOIN Transportadores ON Despacho.Transportador = Transportadores.ID) INNER JOIN Clientes ON Despacho.Cliente = Clientes.ID WHERE Despacho.Tipo = 'Madera Rolliza' ORDER BY Despacho.ID DESC;");
            }
            else if (Variables.tipo == 4)
            {
                if (ano != -1)
                    Variables.cargar(data, "SELECT Despacho.ID,Despacho.despacho, Clientes.cliente, (Transportadores.nombres + ' ' + Transportadores.apellidos) As Transportador, numICA, numICAWeb, FSC, Fecha, Hora, Tipo FROM (Despacho INNER JOIN Transportadores ON Despacho.Transportador = Transportadores.ID) INNER JOIN Clientes ON Despacho.Cliente = Clientes.ID WHERE (((Year([Fecha]))=" + ano + ")) AND Despacho.Tipo = 'Otro' ORDER BY Despacho.ID DESC;");
                else
                    Variables.cargar(data, "SELECT Despacho.ID,Despacho.despacho, Clientes.cliente, Transportadores.nombres, numICA, numICAWeb, FSC, Fecha, Hora, Tipo FROM (Despacho INNER JOIN Transportadores ON Despacho.Transportador = Transportadores.ID) INNER JOIN Clientes ON Despacho.Cliente = Clientes.ID WHERE Despacho.Tipo = 'Otro' ORDER BY Despacho.ID DESC;");
            }
            else
            {
                if (ano != -1)
                    Variables.cargar(data, "SELECT Despacho.ID,Despacho.despacho, Clientes.cliente, (Transportadores.nombres + ' ' + Transportadores.apellidos) As Transportador, numICA, numICAWeb, FSC, Fecha, Hora, Tipo FROM (Despacho INNER JOIN Transportadores ON Despacho.Transportador = Transportadores.ID) INNER JOIN Clientes ON Despacho.Cliente = Clientes.ID WHERE (((Year([Fecha]))=" + ano + "))ORDER BY Despacho.ID DESC;");
                else
                    Variables.cargar(data, "SELECT Despacho.ID,Despacho.despacho, Clientes.cliente, Transportadores.nombres, numICA, numICAWeb, FSC, Fecha, Hora, Tipo FROM (Despacho INNER JOIN Transportadores ON Despacho.Transportador = Transportadores.ID) INNER JOIN Clientes ON Despacho.Cliente = Clientes.ID ORDER BY Despacho.ID DESC;");
            }
            data.Columns[4].HeaderText = "# ICA";
            data.Columns[5].HeaderText = "# ICA Web";
            data.Columns[1].DefaultCellStyle.Font = new Font(data.DefaultCellStyle.Font, FontStyle.Underline);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmDespachoOpcion newFrm = new frmDespachoOpcion();
            this.Hide();
            newFrm.Show();
            this.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 1)
            {
                string tipo = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[dataGridView2.Columns.Count - 1].Value.ToString();
                int tipo2 = 0;
                if (tipo.Equals("Otro"))
                    tipo2 = 1;
                frmCrearDespacho newFrm = new frmCrearDespacho(tipo2, Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString()));
                this.Hide();
                newFrm.Show();
                this.Close();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
                cargar(dataGridView1, DateTime.Now.Year - 1);
            else if (tabControl1.SelectedIndex == 2)
                cargar(dataGridView3, -1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmFacturacion newFrm = new frmFacturacion();
            newFrm.Show();
        }
    }
}
