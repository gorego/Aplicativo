using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Aplicativo
{
    public partial class frmRecibos : Form
    {
        public frmRecibos()
        {
            InitializeComponent();
            Variables.cargar(dataGridView2, "SELECT Recibo.*,Lotes.Lote FROM Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote WHERE Recibo.volumenActual > 0 ORDER BY Recibo.ID Desc");
            Variables.cargar(dataGridView1, "SELECT Recibo.*,Lotes.Lote FROM Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote ORDER BY Recibo.ID Desc");
            Variables.cargar(dataGridView3, "SELECT Recibo.*,Lotes.Lote FROM Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote WHERE Recibo.Especie = 'Melina' ORDER BY Recibo.ID Desc");
            Variables.cargar(dataGridView4, "SELECT Recibo.*,Lotes.Lote FROM Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote WHERE Recibo.Especie = 'Teca' ORDER BY Recibo.ID Desc");
            Variables.cargar(dataGridView5, "SELECT REcibo.*,Lotes.Lote FROM Lotes INNER JOIN Recibo ON Lotes.Codigo = Recibo.Lote WHERE Recibo.Especie <> 'Melina' AND Recibo.Especie <> 'Teca' ORDER BY Recibo.ID Desc");
            crearFormatoData(dataGridView1);
            crearFormatoData(dataGridView2);
            crearFormatoData(dataGridView3);
            crearFormatoData(dataGridView4);
            crearFormatoData(dataGridView5);
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
        }

    }
}
