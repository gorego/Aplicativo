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
    public partial class frmBodega : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();

        public frmBodega()
        {
            InitializeComponent();
            Variables.cargar(dataGridView1, "SELECT Productos.ID, Productos.Codigo, Productos.Especie, COUNT(Paquete.ID), (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)), ((((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)) * COUNT(Paquete.ID)) FROM Productos INNER JOIN Paquete ON Productos.ID = Paquete.Producto WHERE Paquete.volumenActual > 0 AND Productos.Caracteristica = 'Dimensionado' AND Paquete.porcentaje = 1 GROUP BY Productos.ID, Productos.Codigo, Productos.Especie, (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp))");
            generarFormato(dataGridView1);
            Variables.cargar(dataGridView2, "SELECT Productos.ID, Productos.Codigo, Productos.Especie, COUNT(Paquete.ID), (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)), ((((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)) * COUNT(Paquete.ID)) FROM Productos INNER JOIN Paquete ON Productos.ID = Paquete.Producto WHERE Paquete.volumenActual > 0 AND Productos.Caracteristica = 'Dimensionado para Secar' AND Paquete.porcentaje = 1 GROUP BY Productos.ID, Productos.Codigo, Productos.Especie, (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp))");
            generarFormato(dataGridView2);
            Variables.cargar(dataGridView3, "SELECT Productos.ID, Productos.Codigo, Productos.Especie, COUNT(Paquete.ID), (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)), ((((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)) * COUNT(Paquete.ID)) FROM Productos INNER JOIN Paquete ON Productos.ID = Paquete.Producto WHERE Paquete.volumenActual > 0 AND Productos.Caracteristica = 'Seco Dimensionado' AND Paquete.porcentaje = 1 GROUP BY Productos.ID, Productos.Codigo, Productos.Especie, (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp))");
            generarFormato(dataGridView3);
            Variables.cargar(dataGridView4, "SELECT Productos.ID, Productos.Codigo, Productos.Especie, COUNT(Paquete.ID), (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)), ((((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)) * COUNT(Paquete.ID)) FROM Productos INNER JOIN Paquete ON Productos.ID = Paquete.Producto WHERE Paquete.volumenActual > 0 AND Productos.Caracteristica = 'Seco Cepillado' AND Paquete.porcentaje = 1 GROUP BY Productos.ID, Productos.Codigo, Productos.Especie, (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp))");
            generarFormato(dataGridView4);
            Variables.cargar(dataGridView5, "SELECT Productos.ID, Productos.Codigo, Productos.Especie, COUNT(Paquete.ID), (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)), ((((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp)) * COUNT(Paquete.ID)) FROM Productos INNER JOIN Paquete ON Productos.ID = Paquete.Producto WHERE Paquete.volumenActual > 0 AND Productos.Caracteristica = 'Slats o Lamelas' AND Paquete.porcentaje = 1 GROUP BY Productos.ID, Productos.Codigo, Productos.Especie, (((Productos.anchoProd * Productos.altoProd * Productos.largoProd)/1000000000) * (Productos.numAnchoEmp* Productos.numAltoEmp))");
            generarFormato(dataGridView5);

            label2.Text = "Total: " + getTotal(dataGridView1);
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView2.Columns[1].DefaultCellStyle.Font = new Font(dataGridView2.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView3.Columns[1].DefaultCellStyle.Font = new Font(dataGridView3.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView4.Columns[1].DefaultCellStyle.Font = new Font(dataGridView4.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView5.Columns[1].DefaultCellStyle.Font = new Font(dataGridView5.DefaultCellStyle.Font, FontStyle.Underline);
            dataGridView1.Columns[0].Visible = false;
        }

        public void generarFormato(DataGridView data) {
            data.Columns[1].HeaderText = "Producto";
            data.Columns[3].HeaderText = "# Paquetes";
            data.Columns[4].HeaderText = "Vol. Paquete";
            data.Columns[5].HeaderText = "Vol. Total";

            for (int i = 0; i < data.Rows.Count; i++)
            {
                data.Rows[i].Cells[5].Value = Math.Round(double.Parse(data.Rows[i].Cells[5].Value.ToString()), 4, MidpointRounding.AwayFromZero);
            }
            data.Columns[0].Visible = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Especie LIKE '%{0}%'", textBox1.Text);
                label2.Text = "Total: " + getTotal(dataGridView1) + " m3";
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format("Especie LIKE '%{0}%'", textBox1.Text);
                label2.Text = "Total: " + getTotal(dataGridView2) + " m3";
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = string.Format("Especie LIKE '%{0}%'", textBox1.Text);
                label2.Text = "Total: " + getTotal(dataGridView3) + " m3";
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                (dataGridView4.DataSource as DataTable).DefaultView.RowFilter = string.Format("Especie LIKE '%{0}%'", textBox1.Text);
                label2.Text = "Total: " + getTotal(dataGridView4) + " m3";
            }
            else
            {
                (dataGridView5.DataSource as DataTable).DefaultView.RowFilter = string.Format("Especie LIKE '%{0}%'", textBox1.Text);
                label2.Text = "Total: " + getTotal(dataGridView5) + " m3";
            }
        }

        public string getTotal(DataGridView data)
        {
            string total = "";
            double valor = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                valor += double.Parse(data.Rows[i].Cells[5].Value.ToString());
            }
            total = valor.ToString();
            return total;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(tabControl1.SelectedIndex == 0)
                label2.Text = "Total: " + getTotal(dataGridView1);
            else if (tabControl1.SelectedIndex == 1)
                label2.Text = "Total: " + getTotal(dataGridView2);
            else if (tabControl1.SelectedIndex == 2)
                label2.Text = "Total: " + getTotal(dataGridView3);
            else if (tabControl1.SelectedIndex == 3)
                label2.Text = "Total: " + getTotal(dataGridView4);
            else if (tabControl1.SelectedIndex == 4)
                label2.Text = "Total: " + getTotal(dataGridView5);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                frmPaqueteOP newFrm = new frmPaqueteOP(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex == 1)
            {
                frmPaqueteOP newFrm = new frmPaqueteOP(dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView3.CurrentCell.ColumnIndex == 1)
            {
                frmPaqueteOP newFrm = new frmPaqueteOP(dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView4.CurrentCell.ColumnIndex == 1)
            {
                frmPaqueteOP newFrm = new frmPaqueteOP(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView5.CurrentCell.ColumnIndex == 1)
            {
                frmPaqueteOP newFrm = new frmPaqueteOP(dataGridView5.Rows[dataGridView5.CurrentCell.RowIndex].Cells[0].Value.ToString());
                newFrm.Show();
            }
        }

    }
}
