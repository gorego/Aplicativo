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
using System.Globalization;
using System.IO;

namespace Aplicativo
{
    public partial class frmOrdenFormatos : Form
    {
        String connectionString = Variables.connectionString;
        OleDbConnection conn = new OleDbConnection();
        string area, actividad, unidad, supervisor, estado, lote, predio, OT, fechaInicial, fechaFinal, descripcion, especie, estadoFSC, FSC, userName, nomOrden;
        int semanaOrden, semanaActual, codigo, transportador, cliente,semillas, transp, operador;
        int termino = 0;
        int semanaFinal = 0;
        int tipousuario;
        string nomOperador, NIT = "";
        bool sw = false, sw2 = true;

        string[,] ADF003 = new string[,] { { "Motosierras", "Gasolina", "Aceite 2T", "Aceite de Cadena", "Cadenas", "Limas", "Bomba de Espalda", "Quimico", "Machetes", "Guantes", "Lentes", "Tractor + Trailer Sencillo", "Tractor +  Trailer Forestal" }, { "Unidad", "Litros", "Litros", "Litros", "Unidad", "Unidad", "Unidad", "Litros", "Unidad", "Unidad", "Unidad", "Unidad", "Unidad" } };
        string[,] ADF004 = new string[,] { { "Cosechador", "Motosierra", "Gasolina", "Aceite 2T", "Aceite de Cadena", "Cadenas", "Limas", "Bomba de Espalda", "Quimico", "Machetes", "Guantes", "Lentes", "Tractor + Trailer Sencillo", "Tractor +  Trailer Forestal" }, { "Unidad", "Unidad", "Litros", "Litros", "Litros", "Unidad", "Unidad", "Unidad", "Litros", "Unidad", "Unidad", "Unidad", "Unidad", "Unidad" } };
        string[,] ADF005 = new string[,] { { "Combustible", "Aceite Hyd", "Aceite Motor", "Aceite de Cadena", "Cadenas", "Espadas", "Limas", "Machetes", "Guantes" }, { "Gal", "Litros", "Litros", "Litros", "Unidad", "Unidad", "Unidad", "Unidad", "Unidad" } };
        string[,] ADF006 = new string[,] { { "Tractor", "Combustible", "Aceite Hyd - Caja", "Aceite Motor" }, { "Horas", "Gal", "Litros", "Litros" } };
        string[] Reporte = new string[] { "Mangueras", "Filtro Combustible", "Llantas", "Otro" };
        string[,] ADF009 = new string[,] { { "Semilla Melina", "Semilla Teca", "Semilla Eucalipto", "Semilla Otra", "Sacos", "Guantes", "Machetes", "Lentes" }, { "Kg", "Kg", "Kg", "Kg", "Unidad", "Unidad", "Unidad", "Unidad" } };
        string[] mantMotor = new string[] { "Cambio Aceite", "Cambio Filtro Aceite", "Cambio Liquido Refrigerante", "Cambio Filtro Refrigerante", "Cambio Filtro Primario Aire", "Cambio Filtro Secundario Aire", "Cambio Separador de Agua", "Cambio Filtro Combustible" };
        string[] mantDiferenciales = new string[] { "Cambio Aceite Caja de Bombas", "Cambio Aceite Transmision", "Cambio Aceite Diferencial Del", "Cambio Aceite Diferencial Tra" };
        string[] Cliente = new string[5];
        int[] Maquinarias = new int[101];
        int[] empleadosSemillas = new int[101];
        int[] especiesSemillas = new int[101];


        public frmOrdenFormatos(string orden, int tipo, int usuario)
        {            
            InitializeComponent();
            tipousuario = usuario;
            if (tipousuario == 1)
            {
                button21.Enabled = false;
            }
            userName = Variables.userName;
            OT = orden;
            hideFormatos();
            getFormatos(orden);
            getOrden(orden);
            getLote(codigo);
            getOperador(operador);
            this.Text = "Formatos de la Orden #" + nomOrden;
            eliminarFormatos();
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd/MM/yyyy";
            if (sw2)
            {
                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                IFormatProvider culture = new System.Globalization.CultureInfo("es-CO", true);
                DateTime date = DateTime.Parse(fechaInicial, culture, System.Globalization.DateTimeStyles.AssumeLocal);
                Calendar cal = dfi.Calendar;
                DateTime date1 = DateTime.Now;
                int yearOrden = date.Year;
                int yearActual = date1.Year;
                int yearDif = (yearActual - yearOrden) * 52;
                semanaOrden = cal.GetWeekOfYear(date, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
                int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek) - semanaOrden + yearDif;
                semanaActual = tipo;
                semanaFinal = semanaActual - semanaOrden + yearDif;
                if (semanaFinal <= 0)
                {
                    linkLabel1.Visible = false;
                }
                if (semana == semanaFinal)
                {
                    linkLabel2.Visible = false;
                }
                label104.Text = "Semana #: " + (semanaFinal + 1);
                //if (tipo == 1)
                //    semanaActual = semanaActual - 1;
                //cargar(comboBox8, "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias", "Maquina");                
                //cargar(comboBox10, "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias", "Maquina");
                //cargar(comboBox9, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                //cargar(comboBox11, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                for (int i = 0; i < tabControl1.TabPages.Count; i++)
                {
                    if (tabControl1.TabPages[i].Text.Equals("ADF002"))
                    {
                        cargarEmpleados(orden);
                        cargarSemanal(semanaFinal, orden);
                        if (!ADF002Existe(semanaFinal, orden))
                        {
                            crearADF002(semanaFinal, orden);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF003"))
                    {
                        crearADFInsumos(dataGridView1, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo",ADF003);                        
                        cargarADF(semanaFinal, orden, "ADF003", dataGridView1);
                        if (!ADFExiste(semanaFinal, orden, "ADF003"))
                        {
                            crearADF(semanaFinal, orden, "ADF003", dataGridView1);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF004"))
                    {
                        crearADFInsumos(dataGridView2, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo", ADF004);                        
                        //formato(dataGridView2, ADF004);
                        cargarADF(semanaFinal, orden, "ADF004", dataGridView2);
                        if (!ADFExiste(semanaFinal, orden, "ADF004"))
                        {
                            crearADF(semanaFinal, orden, "ADF004", dataGridView2);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF005"))
                    {
                        crearADFInsumos(dataGridView4, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo", ADF005);
                        int numMaquinas = 0;
                        getInfo(Int32.Parse(OT), Maquinarias, numMaquinas);
                        //formato(dataGridView4, ADF005);
                        formato(dataGridView6, Reporte);
                        formato(dataGridView5, unidad);
                        dataGridView5.Rows[0].Cells[3].Value = Maquinarias[0];
                        dataGridView5.Rows[0].Cells[0].Value = 1;
                        cargarADF(semanaFinal, orden, "ADF005", dataGridView4);
                        cargarADF006(semanaFinal, orden, "ADF005-2", dataGridView5);
                        cargarDaños(semanaFinal, orden, dataGridView6, "ADF005");
                        if (!ADFExiste(semanaFinal, orden, "ADF005"))
                        {
                            crearADF(semanaFinal, orden, "ADF005", dataGridView4);
                        }
                        if (!ADFExiste(semanaFinal, orden, "ADF005-2"))
                        {
                            crearADF(semanaFinal, orden, "ADF005-2", dataGridView5);
                        }
                        if (!dañosExiste(semanaFinal, orden, "ADF005"))
                        {
                            crearDaños(semanaFinal, orden, "ADF005", dataGridView6);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF006"))
                    {                        
                        crearADF006(dataGridView9, "SELECT * FROM Insumos WHERE Descripcion LIKE '%", "Modelo", ADF006);
                        int numMaquinas = 0;
                        getInfo(Int32.Parse(OT), Maquinarias, numMaquinas);
                        dataGridView9.Rows[0].Cells[3].Value = Maquinarias[0];
                        //formato(dataGridView9, ADF006);
                        formato(dataGridView7, Reporte);
                        formato(dataGridView8, unidad);
                        dataGridView8.Rows[0].Cells[3].Value = Maquinarias[0];
                        dataGridView8.Rows[0].Cells[0].Value = 1;
                        cargarADF(semanaFinal, orden, "ADF006", dataGridView9);
                        cargarADF006(semanaFinal, orden, "ADF006-2", dataGridView8);
                        cargarDaños(semanaFinal, orden, dataGridView7, "ADF006");
                        if (!ADFExiste(semanaFinal, orden, "ADF006"))
                        {
                            crearADF(semanaFinal, orden, "ADF006", dataGridView9);                            
                        }
                        if (!ADFExiste(semanaFinal, orden, "ADF006-2"))
                        {
                            crearADF(semanaFinal, orden, "ADF006-2", dataGridView8);
                        }
                        if (!dañosExiste(semanaFinal, orden, "ADF006"))
                        {
                            crearDaños(semanaFinal, orden, "ADF006", dataGridView7);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF007"))
                    {
                        crearADF(dataGridView10, "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias WHERE Tipo = 'Motosierra' OR Tipo = 'Cosechador'", "Maquina");
                        formato(dataGridView10, "Arboles", 0);
                        dataGridView10.Rows[0].Cells[0].Value = 1;
                        dataGridView10.Rows[0].Cells[2].Value = "Equipo: ";
                        cargarADFExtraccion(semanaFinal, orden, "ADF007", dataGridView10);
                        if (!ADFExiste(semanaFinal, orden, "ADF007"))
                        {
                            crearADF2(semanaFinal, orden, "ADF007", dataGridView10);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF008"))
                    {
                        crearADF(dataGridView11, "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Placa) As Maquina FROM Maquinarias WHERE Tipo = 'Motosierra' OR Tipo = 'Cosechador'", "Maquina");
                        formato(dataGridView11, "Arboles", 0);
                        dataGridView11.Rows[0].Cells[0].Value = 1;
                        dataGridView11.Rows[0].Cells[2].Value = "Equipo: ";
                        cargarADFExtraccion(semanaFinal, orden, "ADF008", dataGridView11);
                        if (!ADFExiste(semanaFinal, orden, "ADF008"))
                        {
                            crearADF2(semanaFinal, orden, "ADF008", dataGridView11);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF009"))
                    {
                        crearADF(dataGridView12, "SELECT * FROM Insumos WHERE Clase LIKE '", "Modelo", ADF009);                        
                        //formato(dataGridView12, ADF009);
                        cargarADF(semanaFinal, orden, "ADF009", dataGridView12);
                        if (!ADFExiste(semanaFinal, orden, "ADF009"))
                        {
                            crearADF(semanaFinal, orden, "ADF009", dataGridView12);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF010"))
                    {
                        crearADFSemillas(dataGridView13, "SELECT ID,(Nombres + ' ' + Apellidos) As Nombre FROM Trabajadores", "Nombre");
                        dataGridView13.Rows[dataGridView13.Rows.Count - 1].Cells[0].Value = dataGridView13.Rows.Count;
                        for (int j = 0; j < 7; j++)
                        {
                            dataGridView13.Rows[dataGridView13.Rows.Count - 1].Cells[j + 5].Value = 0;
                        }
                        cargarSemilla(semanaFinal, orden, dataGridView13);
                        for (int j = 0; j < semillas; j++)
                        {
                            dataGridView13.Rows[j].Cells[2].Value = empleadosSemillas[j];
                            dataGridView13.Rows[j].Cells[4].Value = especiesSemillas[j];
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF011"))
                    {
                        Variables.cargar2(comboBox12, "SELECT Clase From Insumos GROUP BY Clase", "Clase");
                        Variables.cargar2(comboBox10, "SELECT Marca From Insumos GROUP BY Marca", "Marca");
                        Variables.cargar(comboBox8, "SELECT * From Insumos", "Modelo");
                        crearADF(dataGridView14, "SELECT ID, (Clase + ' ' + Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");                        
                        cargarADFExtraccion(semanaFinal, orden, "ADF011", dataGridView14);
                        termino = 1;
                        //dataGridView14.Rows[0].Cells[0].Value = 1;
                        if (!ADFExiste(semanaFinal, orden, "ADF011"))
                        {
                            crearADF3(semanaFinal, orden, "ADF011", dataGridView14);
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF012"))
                    {
                        cargar(comboBox1, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                        cargar(comboBox2, "SELECT ID, (Tipo + ' / ' + Marca + '/ ' + Modelo + ' / ' + Placa) As Maquina FROM Maquinarias", "Maquina");
                        comboBox1.SelectedItem = null;
                        comboBox2.SelectedItem = null;
                        formato(material1, 0);
                        formato(material2, 75);
                        formato(material3, 150);
                        formato(material4, 225);
                        if (estadoFSC.Equals("Si"))
                        {
                            radioButton2.Checked = true;
                            textBox1.Text = FSC;
                        }
                        else
                        {
                            radioButton1.Checked = true;
                        }
                        if (especie.Equals("Melina"))
                            radioButton4.Checked = true;
                        else if (especie.Equals("Teca"))
                            radioButton3.Checked = true;
                        else
                        {
                            radioButton31.Checked = true;
                            textBox2.Text = especie;
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF013"))
                    {
                        cargar(comboBox3, "SELECT ID, (Tipo + ' / ' + Marca + '/ ' + Modelo + ' / ' + Placa) As Maquina FROM Maquinarias", "Maquina");
                        cargar(comboBox4, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                        comboBox3.SelectedItem = null;
                        comboBox4.SelectedItem = null;
                        if (estadoFSC.Equals("Si"))
                        {
                            radioButton13.Checked = true;
                            textBox6.Text = FSC;
                        }
                        else
                        {
                            radioButton14.Checked = true;
                        }
                        if (especie.Equals("Melina"))
                            radioButton12.Checked = true;
                        else if (especie.Equals("Teca"))
                            radioButton11.Checked = true;
                        else
                        {
                            radioButton30.Checked = true;
                            textBox5.Text = especie;
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF014"))
                    {
                        cargar(comboBox9, "SELECT ID, (Nombres + ' ' + Apellidos) As nombre FROM Transportadores", "nombre");
                        comboBox9.SelectedItem = null;
                        formato(material5, 0);
                        formato(material6, 75);
                        formato(material7, 150);
                        formato(material8, 225);
                        getInfo(cliente, Cliente);
                        textBox25.Text = Cliente[0];
                        textBox26.Text = Cliente[1];
                        textBox27.Text = Cliente[2];
                        textBox28.Text = Cliente[3];
                        textBox29.Text = Cliente[4];
                        comboBox9.SelectedValue = transportador;
                        if (estadoFSC.Equals("Si"))
                        {
                            radioButton22.Checked = true;
                            textBox24.Text = FSC;
                        }
                        else
                        {
                            radioButton23.Checked = true;
                        }
                        if (especie.Equals("Melina"))
                            radioButton21.Checked = true;
                        else if (especie.Equals("Teca"))
                            radioButton20.Checked = true;
                        else
                        {
                            radioButton29.Checked = true;
                            textBox23.Text = especie;
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF015"))
                    {
                        cargar(comboBox11, "SELECT ID, (Nombres + ' ' + Apellidos) As nombre FROM Transportadores", "nombre");
                        comboBox11.SelectedItem = null;
                        getInfo(cliente, Cliente);
                        textBox34.Text = Cliente[0];
                        textBox33.Text = Cliente[1];
                        textBox32.Text = Cliente[2];
                        textBox31.Text = Cliente[3];
                        textBox30.Text = Cliente[4];
                        comboBox11.SelectedValue = transportador;
                        if (estadoFSC.Equals("Si"))
                        {
                            radioButton19.Checked = true;
                            textBox37.Text = FSC;
                        }
                        else
                        {
                            radioButton24.Checked = true;
                        }
                        if (especie.Equals("Melina"))
                            radioButton18.Checked = true;
                        else if (especie.Equals("Teca"))
                            radioButton17.Checked = true;
                        else
                        {
                            radioButton28.Checked = true;
                            textBox36.Text = especie;
                        }
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF016"))
                    {
                        cargar(comboBox5, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                        comboBox5.SelectedItem = null;
                        //formatoMant(dataGridView15, mantMotor);
                        int numMaquinas = 0;
                        getInfo(Int32.Parse(OT), Maquinarias, numMaquinas);
                        textBox11.Text = getHorometro(Maquinarias[0]).ToString();
                        comboBox5.SelectedValue = getEmpleado(Int32.Parse(OT));
                        string[] mantHidraulico = new string[] { "Cambio Aceite", "Cambio Filtro Aceite 1", "Cambio Filtro Aceite 2", "Cambio Filtro Aceite 3", "Cambio Filtro Aceite 4", "Cambio Separador de Agua" };
                        crearADFMant(dataGridView15, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantMotor);
                        crearADFMant(dataGridView16, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantHidraulico);
                        crearADFMant(dataGridView17, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantDiferenciales);
                        crearADFMantDetalle(dataGridView18, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");
                        cargarADFMant(semanaFinal, OT, "ADF016", dataGridView15,"Motor");
                        cargarADFMant(semanaFinal, OT, "ADF016", dataGridView16,"Hidraulico");
                        cargarADFMant(semanaFinal, OT, "ADF016", dataGridView17,"Diferenciales");
                        cargarADFMantDetalle(semanaFinal, OT, "ADF016", dataGridView18,comboBox5,dateTimePicker1);
                        //formatoMant(dataGridView16, mantHidraulico);
                        //formatoMant(dataGridView17, mantDiferenciales);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF017"))
                    {
                        cargar(comboBox6, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                        comboBox6.SelectedItem = null;
                        //formatoMant(dataGridView22, mantMotor);
                        int numMaquinas = 0;
                        getInfo(Int32.Parse(OT), Maquinarias, numMaquinas);
                        textBox18.Text = getHorometro(Maquinarias[0]).ToString();
                        comboBox6.SelectedValue = getEmpleado(Int32.Parse(OT));
                        string[] mantHidraulico = new string[] { "Cambio Aceite", "Cambio Filtro Tanque", "Cambio Filtro Transmisión" };
                        crearADFMant(dataGridView22, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantMotor);
                        crearADFMant(dataGridView21, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantHidraulico);
                        crearADFMant(dataGridView20, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantDiferenciales);
                        crearADFMantDetalle(dataGridView19, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos", "Insumo");
                        cargarADFMant(semanaFinal, OT, "ADF017", dataGridView22,"Motor");
                        cargarADFMant(semanaFinal, OT, "ADF017", dataGridView21,"Hidraulico");
                        cargarADFMant(semanaFinal, OT, "ADF017", dataGridView20,"Diferenciales");
                        cargarADFMantDetalle(semanaFinal, OT, "ADF017", dataGridView19, comboBox6, dateTimePicker2);
                        //formatoMant(dataGridView21, mantHidraulico);
                        //formatoMant(dataGridView20, mantDiferenciales);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF018"))
                    {
                        cargar(comboBox7, "SELECT ID, (Nombres + '  ' + Apellidos) As nombre FROM Trabajadores", "nombre");
                        comboBox7.SelectedItem = null;
                        //formatoMant(dataGridView26, mantMotor);
                        int numMaquinas = 0;
                        getInfo(Int32.Parse(OT), Maquinarias, numMaquinas);
                        textBox21.Text = getHorometro(Maquinarias[0]).ToString();
                        comboBox7.SelectedValue = getEmpleado(Int32.Parse(OT));
                        string[] mantHidraulico = new string[] { "Cambio Aceite", "Cambio Filtro Hidaulico", "Otro" };
                        crearADFMant(dataGridView26, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantMotor);
                        crearADFMant(dataGridView25, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantHidraulico);
                        crearADFMant(dataGridView24, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo", mantDiferenciales);
                        crearADFMantDetalle(dataGridView23, "SELECT ID,(Marca + ' ' + Modelo) As Insumo FROM Insumos WHERE Clase LIKE '%Aceite%' OR Clase LIKE '%Filtro%'", "Insumo");
                        cargarADFMant(semanaFinal, OT, "ADF018", dataGridView26,"Motor");
                        cargarADFMant(semanaFinal, OT, "ADF018", dataGridView25,"Hidraulico");
                        cargarADFMant(semanaFinal, OT, "ADF018", dataGridView24,"Diferenciales");
                        cargarADFMantDetalle(semanaFinal, OT, "ADF018", dataGridView23, comboBox7, dateTimePicker3);
                        //formatoMant(dataGridView25, mantHidraulico);
                        //formatoMant(dataGridView24, mantDiferenciales);
                    }
                    else if (tabControl1.TabPages[i].Text.Equals("ADF019"))
                    {

                    }
                }
            }
        }

        public void cargarADFMant(int semana, string orden, string adf, DataGridView data, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (data.Rows[i].Cells[1].Value != null)
                {
                    string query = "SELECT * FROM formatoMantenimiento WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND ADF = '" + adf + "' AND Tipo = '" + tipo + "'";
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
                            if (!myReader.IsDBNull(4))
                                if (myReader.GetInt32(4) != 0)
                                    data.Rows[i].Cells[2].Value = myReader.GetInt32(6);
                            data.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                            data.Rows[i].Cells[3].Value = myReader.GetInt32(7);
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
            }
        }

        public void cargarADFMantDetalle(int semana, string orden, string adf, DataGridView data,ComboBox empleado, DateTimePicker fecha)
        {
            string query = "SELECT * FROM formatoMantenimiento WHERE Tipo = 'Adicional' AND Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "'";
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
                    data.Rows.Add();
                    data.Rows[i].Cells[1].Value = myReader.GetString(5);
                    if (!myReader.IsDBNull(4))
                        if (myReader.GetInt32(4) != 0)
                            data.Rows[i].Cells[2].Value = myReader.GetInt32(6);
                    data.Rows[i].Cells[0].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[3].Value = myReader.GetInt32(7);
                    empleado.SelectedValue = myReader.GetInt32(12);
                    fecha.Value = DateTime.ParseExact(myReader.GetString(3), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
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


        public void agregarLog(string accion, string usuario)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO historicoIngresos (Usuario,Accion,Fecha) VALUES (@Usuario,@Accion,@Fecha)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Usuario", OleDbType.VarChar).Value = usuario;
                cmd.Parameters.Add("@Accion", OleDbType.VarChar).Value = accion;
                cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year + " - " + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" + DateTime.Now.Second;
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

        public void crearADFSemillas(DataGridView data, string query, string display)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Nombre";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            cargarDetalle(combo, query, display, data);
            data.Columns.Add(combo);
            data.Columns.Add("Column3", "Cedula");
            combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Especie Recogida";
            cargarDetalle(combo, "SELECT ID, (Clase + ' ' + Modelo) As insumo FROM Insumos WHERE Clase LIKE '%Semilla%'", "insumo", data);
            data.Columns.Add(combo);
            data.Columns.Add("Column4", "Lunes");
            data.Columns.Add("Column5", "Martes");
            data.Columns.Add("Column6", "Miercoles");
            data.Columns.Add("Column7", "Jueves");
            data.Columns.Add("Column8", "Viernes");
            data.Columns.Add("Column9", "Sabado");
            data.Columns.Add("Column10", "Total");
            data.Columns[0].FillWeight = 40;            
            data.Columns[2].FillWeight = 300;
            data.Columns[3].FillWeight = 150;
            data.Columns[4].FillWeight = 150;
        }


        public void crearADF(DataGridView data, string query, string display)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            cargarDetalle(combo, query, display, data);
            data.Columns.Add(combo);
            for (int i = 0; i < 6; i++)
            {
                data.Columns.Add("Column" + i + 4, "Unidad");
                data.Columns[4 + (i * 2)].ReadOnly = true;
                data.Columns.Add("Column" + i + 5, "Cantidad");
            }
            data.Columns.Add("Column20", "Total");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 100;
            data.Columns[3].FillWeight = 300;
        }

        public void crearADF(DataGridView data, string query, string display,string[,] formatos)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            data.Columns.Add(combo);
            for (int i = 0; i < 6; i++)
            {
                data.Columns.Add("Column" + i + 4, "Unidad");
                data.Columns[4 + (i * 2)].ReadOnly = true;
                data.Columns.Add("Column" + i + 5, "Cantidad");
            }
            data.Columns.Add("Column20", "Total");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 200;
            data.Columns[3].FillWeight = 300;
            formato(data, formatos);            
            cargarDetalle(query, display, data,formatos);
        }

        public void crearADFInsumos(DataGridView data, string query, string display, string[,] formatos)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            data.Columns.Add(combo);
            for (int i = 0; i < 6; i++)
            {
                data.Columns.Add("Column" + i + 4, "Unidad");
                data.Columns[4 + (i * 2)].ReadOnly = true;
                data.Columns.Add("Column" + i + 5, "Cantidad");
            }
            data.Columns.Add("Column20", "Total");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 200;
            data.Columns[3].FillWeight = 300;
            formato(data, formatos);
            cargarDetalleInsumos(query, display, data, formatos);
        }


        public void crearADFMant(DataGridView data, string query, string display, string[] formatos)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Insumo";            
            data.Columns.Add("Column2", "ID");
            data.Columns[0].Visible = false;
            data.Columns.Add("Column3", "Detalle");            
            data.Columns.Add(combo);
            data.Columns.Add("Column20", "Cantidad");
            data.Columns[1].FillWeight = 50;
            data.Columns[2].FillWeight = 50;
            data.Columns[3].FillWeight = 25;
            formatoMant(data, formatos);
            cargarDetalleMant(query, display, data, formatos);
        }

        public void crearADFMantDetalle(DataGridView data, string query, string display)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Insumo";
            data.Columns.Add("Column2", "ID");
            data.Columns[0].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            data.Columns.Add(combo);
            data.Columns.Add("Column20", "Cantidad");
            data.Columns[1].FillWeight = 50;
            data.Columns[2].FillWeight = 50;
            data.Columns[3].FillWeight = 25;            
            cargarDetalle(combo,query, display, data);
        }

        public void crearADF006(DataGridView data, string query, string display, string[,] formatos)
        {
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.HeaderText = "Modelo";
            data.Columns.Add("Column1", "#");
            data.Columns.Add("Column2", "ID");
            data.Columns[1].Visible = false;
            data.Columns.Add("Column3", "Detalle");
            data.Columns.Add(combo);
            for (int i = 0; i < 6; i++)
            {
                data.Columns.Add("Column" + i + 4, "Unidad");
                data.Columns[4 + (i * 2)].ReadOnly = true;
                data.Columns.Add("Column" + i + 5, "Cantidad");
            }
            data.Columns.Add("Column20", "Total");
            data.Columns[0].FillWeight = 40;
            data.Columns[2].FillWeight = 200;
            data.Columns[3].FillWeight = 350;
            formato(data, formatos);
            cargarDetalleADF006(query, display, data, formatos);
        }

        public void cargar(ComboBox combo, string query, string display)
        {
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = display;
            combo.ValueMember = "ID";
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void cargarDetalle(DataGridViewComboBoxColumn combo, string query, string display, DataGridView data)
        {
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = display;
            combo.ValueMember = "ID";
        }

        public void cargarDetalle(string query, string display, DataGridView data,string[,] formato)
        {
            for (int i = 0; i < (formato.Length) / 2; i++)
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)(data.Rows[i].Cells[3]);
                //Ejecutar el query y llenar el ComboBox.
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand(query + formato[0,i] + "'", conn);
                DataTable maquinaria = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                combo.DataSource = ds.Tables[0];
                combo.DisplayMember = display;
                combo.ValueMember = "ID";
            }
        }

        public void cargarDetalleInsumos(string query, string display, DataGridView data, string[,] formato)
        {
            for (int i = 0; i < (formato.Length) / 2; i++)
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)(data.Rows[i].Cells[3]);
                //Ejecutar el query y llenar el ComboBox.
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand(query + formato[0, i] + "%'", conn);
                DataTable maquinaria = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                combo.DataSource = ds.Tables[0];
                combo.DisplayMember = display;
                combo.ValueMember = "ID";
            }
        }

        public void cargarDetalleMant(string query, string display, DataGridView data, string[] formato)
        {
            for (int i = 0; i < (formato.Length); i++)
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)(data.Rows[i].Cells[2]);
                //Ejecutar el query y llenar el ComboBox.
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand(query, conn);
                DataTable maquinaria = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                combo.DataSource = ds.Tables[0];
                combo.DisplayMember = display;
                combo.ValueMember = "ID";
            }
        }

        public void cargarDetalleADF006(string query, string display, DataGridView data, string[,] formato)
        {
            DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)(data.Rows[0].Cells[3]);
            //Ejecutar el query y llenar el ComboBox.
            conn.ConnectionString = connectionString;
            string query2 = "SELECT ID, (Tipo + ' / ' + Marca + ' / ' + Modelo + ' / ' + Placa) As Maquina FROM Maquinarias";
            OleDbCommand cmd = new OleDbCommand(query2, conn);
            DataTable maquinaria = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            combo.DataSource = ds.Tables[0];
            combo.DisplayMember = "Maquina";
            combo.ValueMember = "ID";
            for (int i = 1; i < (formato.Length) / 2; i++)
            {                
                combo = (DataGridViewComboBoxCell)(data.Rows[i].Cells[3]);
                //Ejecutar el query y llenar el ComboBox.
                conn.ConnectionString = connectionString;
                cmd = new OleDbCommand(query + formato[0, i] + "%'", conn);
                maquinaria = new DataTable();
                da = new OleDbDataAdapter(cmd);
                ds = new DataSet();
                da.Fill(ds);
                combo.DataSource = ds.Tables[0];
                combo.DisplayMember = display;
                combo.ValueMember = "ID";
            }
        }

        public void cargarSemanal(int semana, string orden)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                string query = "SELECT * FROM adf002 WHERE Semana = " + semana + " AND Orden = " + orden + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value.ToString();
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
                        for (int j = 5; j < 11; j++)
                        {
                            if (myReader.GetInt32(j) == 1)
                                dataGridView3.Rows[i].Cells[j].Value = true;
                        }
                        if (myReader.GetInt32(13) == 1)
                            sw2 = true;
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
        }

        public void cargarDaños(int semana, string orden, DataGridView data, string ADF)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM Daños WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND ADF = '" + ADF + "'";
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
                        for (int j = 2; j < 8; j++)
                        {
                            if (myReader.GetInt32(j+2) == 1)
                                data.Rows[i].Cells[j].Value = true;
                        }
                        data.Rows[i].Cells[8].Value = myReader.GetString(11);
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
        }


        public void modificarADF002(int semana, string orden)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE adf002 SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Domingo=@Domingo WHERE Semana = " + semana + " AND Orden = " + orden + " AND Trabajador = " + dataGridView3.Rows[i].Cells[1].Value.ToString());
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[8];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[9];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[10];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[11];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 0;
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

        public void eliminarDias(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores AS t INNER JOIN adf002 AS f ON t.ID = f.Trabajador SET t.diasLaborados = (t.diasLaborados - f.Lunes - f.Martes - f.Miercoles - f.Jueves - f.Viernes - f.Sabado - f.Domingo),f.Estado=1, f.Editable=0 WHERE f.Semana = @semana AND f.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        public int getCostoJornalPasado(int semana, string orden)
        {
            int costo = 0;
            string query = "SELECT (SUM(f.Lunes + f.Martes + f.Miercoles + f.Jueves + f.Viernes + f.Sabado + f.Domingo) * (c.Salario/30)) FROM adf002 AS f INNER JOIN (CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo) ON f.Trabajador = t.ID WHERE f.Orden = " + orden + " AND f.Semana = " + semana + " GROUP BY Salario, Semana, Orden;";
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
                    costo = Int32.Parse(myReader.GetValue(0).ToString());
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return costo;
        }
        public void eliminarCostoJornal(int semana, string orden)
        {
            int costo = getCostoJornalPasado(semana, orden);
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoOrdenes SET costoJornalFinal = costoJornalFinal - @costo WHERE ID = " + orden);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@costo", OleDbType.VarChar).Value = costo;
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

        public void agregarDias(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Trabajadores AS t INNER JOIN adf002 AS f ON t.ID = f.Trabajador SET t.diasLaborados = (t.diasLaborados + f.Lunes + f.Martes + f.Miercoles + f.Jueves + f.Viernes + f.Sabado + f.Domingo),f.Estado=1, f.Editable=0 WHERE f.Semana = @semana AND f.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        public void agregarADF2(int semana, string orden, string ADF)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN Control AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock + c.Lunes + c.Martes + c.Miercoles + c.Jueves + c.Viernes + c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@adf", OleDbType.VarChar).Value = ADF;
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

        public void agregarADF(int semana, string orden, string ADF)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;            
            if(!ADF.Equals("ADF006"))
                cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN Control AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock + c.Lunes + c.Martes + c.Miercoles + c.Jueves + c.Viernes + c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf");
            else
                cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN Control AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock + c.Lunes + c.Martes + c.Miercoles + c.Jueves + c.Viernes + c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf AND c.Detalle <> 'Tractor'");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@adf", OleDbType.VarChar).Value = ADF;
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

        public void agregarADF(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN formatoSemilla AS c ON i.ID = c.Semilla SET i.Cantidad_Stock = (i.Cantidad_Stock + c.Lunes + c.Martes + c.Miercoles + c.Jueves + c.Viernes + c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        public void agregarADFMant(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN formatoMantenimiento AS c ON i.ID = c.Insumo SET i.Cantidad_Stock = (i.Cantidad_Stock + c.Cantidad),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        public void eliminarADFMant(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN formatoMantenimiento AS c ON i.ID = c.Insumo SET i.Cantidad_Stock = (i.Cantidad_Stock - c.Cantidad),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        public void agregarADFEquipo(int semana, string orden, string ADF)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Maquinarias AS m INNER JOIN Control AS c ON m.ID = c.Modelo SET m.Horometro = (m.Horometro + c.Lunes + c.Martes + c.Miercoles + c.Jueves + c.Viernes + c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@adf", OleDbType.VarChar).Value = ADF;
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

        public void agregarVolumen(double diametro, double largo, int cantidad, double volumen, string especie, string trailer, string raleo, string adf)
        {
            volumen = Math.Round(volumen, 3, MidpointRounding.AwayFromZero);
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("INSERT INTO volumenCalculado (Diametro,Largo,Cantidad,Volumen,Especie,Trailer,Raleo,ADF) VALUES (@Diametro,@Largo,@Cantidad,@Volumen,@Especie,@Trailer,@Raleo,@ADF)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = diametro;
                cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = largo;
                cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = cantidad;
                cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = volumen;
                cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = especie;
                cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = trailer;
                cmd.Parameters.Add("@Raleo", OleDbType.VarChar).Value = raleo;
                cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;                
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

        public void eliminarADFEquipo(int semana, string orden, string ADF)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Maquinarias AS m INNER JOIN Control AS c ON m.ID = c.Modelo SET m.Horometro = (m.Horometro - c.Lunes - c.Martes - c.Miercoles - c.Jueves - c.Viernes - c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@adf", OleDbType.VarChar).Value = ADF;
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

        public void eliminarADF(int semana, string orden, string ADF)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            if(!ADF.Equals("ADF006"))
                cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN Control AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock - c.Lunes - c.Martes - c.Miercoles - c.Jueves - c.Viernes - c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf ");
            else
                cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN Control AS c ON i.ID = c.Modelo SET i.Cantidad_Stock = (i.Cantidad_Stock - c.Lunes - c.Martes - c.Miercoles - c.Jueves - c.Viernes - c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden AND c.ADF = @adf AND c.Detalle <> 'Tractor'");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
                cmd.Parameters.Add("@adf", OleDbType.VarChar).Value = ADF;
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

        public void eliminarADF(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Insumos AS i INNER JOIN formatoSemilla AS c ON i.ID = c.Semilla SET i.Cantidad_Stock = (i.Cantidad_Stock - c.Lunes - c.Martes - c.Miercoles - c.Jueves - c.Viernes - c.Sabado),c.Estado=1,c.Editable=0 WHERE c.Semana = @semana AND c.Orden = @orden");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@semana", OleDbType.VarChar).Value = semana;
                cmd.Parameters.Add("@orden", OleDbType.VarChar).Value = orden;
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

        public void agregarCostoJornal(int orden, DataGridView data) {
            int costo = 0;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                costo += (getSalarioDiario(Int32.Parse(data.Rows[i].Cells[1].Value.ToString()))*(Int32.Parse(data.Rows[i].Cells[data.Columns.Count-1].Value.ToString())));
            }

            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE historicoOrdenes SET costoJornalFinal = (costoJornalFinal + @costoJornalFinal) WHERE ID = " + orden);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@costoJornalFinal", OleDbType.VarChar).Value = costo;
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

        public void modificarADF(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd;
                int j = 0;
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1";
                    else
                        adf = "ADF006";
                }
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("UPDATE Control SET Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE Control SET Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'");
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if(j==0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
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
            quitarEditable(semana, orden);
        }

        public void modificarADFEquipos(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count - 1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd;
                int j = 0;
                int k = 0;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    if (data.Rows[i].Cells[1].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE Control SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                        k = 0;
                    }
                    else
                    {
                        cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                        k = 1;
                    }
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE Control SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (k == 1)
                    {
                        cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                        cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                        cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    }
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[5].Value != null)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[7].Value != null)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[9].Value != null)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[11].Value != null)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[13].Value != null)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[15].Value != null)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    if (k == 1)
                    {
                        cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
                    }
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
            quitarEditable(semana, orden);
        }

        public void modificarADF011(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd;
                int j = 0;
                int k = 0;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    if (data.Rows[i].Cells[1].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE Control SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                        k = 0;
                    }
                    else
                    {
                        cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                        k = 1;
                    }
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE Control SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (k == 1)
                    {
                        cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                        cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                        cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    }
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[5].Value != null)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[7].Value != null)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[9].Value != null)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[11].Value != null)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[13].Value != null)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[15].Value != null)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    if (k == 1)
                    {
                        cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
                    }
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
            quitarEditable(semana, orden);
        }


        public void modificarSemilla(int semana, string orden, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count - 1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE formatoSemilla SET Trabajador=@Trabajador,Semilla=@Semilla,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Estado=@Estado,Editable=@Editable WHERE ID = " + data.Rows[i].Cells[1].Value.ToString());
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = "";
                    cmd.Parameters.Add("@Semilla", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[6].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[8].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[10].Value.ToString();
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
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

        public void modificarADF2(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd;
                int j = 0;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    if (data.Rows[i].Cells[1].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE Control SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE ID = " + data.Rows[i].Cells[1].Value);                        
                    }
                    else
                    {
                        cmd = new OleDbCommand("UPDATE Control SET Detalle=@Detalle,Modelo=@Modelo,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado WHERE Orden = " + orden + " AND ADF = '" + adf + "' AND Semana = " + semana);    
                    }
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("UPDATE Control SET Detalle=@Detalle,Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado  WHERE ID = " + data.Rows[i].Cells[1].Value);
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    else
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
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
            quitarEditable(semana, orden);
        }

        public void quitarEditable(int semana, string orden)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Control SET Editable = 0 WHERE Semana = " + semana + " AND Orden = " + orden);
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

        public bool esEditable(string orden, int semana)
        {
            string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Editable = 1 AND Orden = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
                    return true;
                else
                    return false;
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
        }

        public void getInfo(int id, string[] arreglo)
        {
            string query = "SELECT Cliente,Nit,Ciudad,Direccion,Telefono FROM Clientes WHERE ID = " + id;
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
                    for (int i = 0; i < arreglo.Length; i++)
                    {
                        arreglo[i] = myReader.GetValue(i).ToString();   
                    }                    
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

        public void getInfo(int id, int[] array, int numMaquinas)
        {
            string query = "SELECT Maquina FROM ordenMaquinas WHERE Orden = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                numMaquinas = 0;
                while (myReader.Read())
                {
                    array[numMaquinas] = myReader.GetInt32(0);
                    numMaquinas++;
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

        public int getHorometro(int id)
        {
            string query = "SELECT Horometro FROM Maquinarias WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int horometro = 0;
            try
            {
                while (myReader.Read())
                {
                    horometro = myReader.GetInt32(0);                    
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return horometro;
        }

        public int getEmpleado(int id)
        {
            string query = "SELECT Empleado FROM formatoMantenimiento WHERE Orden = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int empleado = 0;
            try
            {
                if (myReader.Read())
                {
                    empleado = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return empleado;
        }

        public string getUnidad(int id)
        {
            string query = "SELECT Unidad_Medida FROM Insumos WHERE ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string jornal = "";
            try
            {
                while (myReader.Read())
                {
                    jornal = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return jornal;
        }

        public int getSalarioDiario(int id)
        {
            string query = "SELECT c.salario FROM CargoLaboral AS c INNER JOIN Trabajadores AS t ON c.ID = t.Cargo WHERE t.ID = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int salario = 0;
            try
            {
                while (myReader.Read())
                {
                    salario = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            if(salario > 0)
                salario = salario / 30; 
            return salario;
        }
        
        public void cargarEmpleados(string orden)
        {
            while (dataGridView3.Rows.Count != 0)
            {
                dataGridView3.Rows.RemoveAt(0);
            }
            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos), t.Cedula, c.Cargo FROM CargoLaboral AS c INNER JOIN (Trabajadores AS t INNER JOIN ordenEmpleados AS s ON t.ID = s.Trabajador) ON c.ID = t.Cargo WHERE s.Orden = " + orden;
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
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[i].Cells[0].Value = i + 1;
                    dataGridView3.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    dataGridView3.Rows[i].Cells[2].Value = myReader.GetString(1);
                    dataGridView3.Rows[i].Cells[3].Value = myReader.GetInt32(2);
                    dataGridView3.Rows[i].Cells[4].Value = myReader.GetString(3);
                    dataGridView3.Rows[i].Cells[12].Value = 0;
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

        public bool ADF002Existe(int semana, string orden)
        {
            string query = "SELECT * FROM adf002 WHERE Semana = " + semana + " AND Orden = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public bool dañosExiste(int semana, string orden, string ADF)
        {
            string query = "SELECT * FROM Daños WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + ADF + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public bool transferenciaExiste(int semana, string orden, string adf, TreeView tree)
        {
            string query = "SELECT * FROM Transferencia WHERE Semana = " + semana + " AND Orden = " + orden + " AND Dia = '" + tree.SelectedNode.Parent.Text + "' AND Recorrido = " + tree.SelectedNode.Text;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public bool ADFExiste(int semana, string orden, string adf)
        {
            string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND adf = '" + adf + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public bool mantExiste(int semana, string orden, string adf)
        {
            string query = "SELECT * FROM formatoMantenimiento WHERE Semana = " + semana + " AND Orden = " + orden + " AND adf = '" + adf + "'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public bool semillaExiste(int semana, string orden)
        {
            string query = "SELECT * FROM formatoSemilla WHERE Semana = " + semana + " AND Orden = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.HasRows)
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

        public void crearADF002(int semana, string orden)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO adf002(Semana,Fecha,Orden,Trabajador,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Domingo,Estado,Editable) VALUES (@Semana,@Fecha,@Orden,@Trabajador,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Domingo,@Estado,@Editable)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = "";
                    cmd.Parameters.Add("@Supervisor", OleDbType.VarChar).Value = orden;
                    cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = dataGridView3.Rows[i].Cells[1].Value.ToString();
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[8];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[9];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[10];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[11];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Domingo", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
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

        public void crearDaños(int semana, string orden, string ADF, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO Daños(Semana,Orden,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,ADF,Descripcion) VALUES (@Semana,@Orden,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@ADF,@Descripcion)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;                    
                    cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value.ToString();
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[2];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[3];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[4];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = ADF;
                    if (data.Rows[i].Cells[8].Value != null)
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = data.Rows[i].Cells[8].Value.ToString();
                    else
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = "";
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

        public void modificarDaños(int semana, string orden, string ADF, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("UPDATE Daños SET Lunes=@Lunes,Martes=@Martes,Miercoles=@Miercoles,Jueves=@Jueves,Viernes=@Viernes,Sabado=@Sabado,Descripcion=@Descripcion WHERE Semana = " + semana + " AND ADF = '" + ADF + "' AND Detalle = '" + data.Rows[i].Cells[1].Value.ToString() + "' AND Orden = " + orden);
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[2];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[3];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[4];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[5];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[6];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    ch1 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[7];
                    if ((bool)ch1.FormattedValue == true)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 1;
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[8].Value != null)
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = data.Rows[i].Cells[8].Value.ToString();
                    else
                        cmd.Parameters.Add("@Descripcion", OleDbType.VarChar).Value = "";
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

        public void cargarADF006(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "'";
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
                        if (!myReader.IsDBNull(4))
                            if (myReader.GetInt32(4) != 0)
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                        data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                        if (!myReader.IsDBNull(5))                            
                            data.Rows[i].Cells[2].Value = myReader.GetString(5);
                        for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
                        }
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
        }

        public void cargarADF(int semana, string orden, string adf, TreeView tree)
        {
            string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "'";
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
                    for (int j = 6, k = 0; j < 12; j++, k++) 
                        for(int i = 0; i < myReader.GetInt32(j); i++)
                            tree.Nodes[k].Nodes.Add((i+1).ToString());
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

        public void cargarADF(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1";
                    else
                        adf = "ADF006";
                }
                string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'";
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
                        if (!myReader.IsDBNull(4))
                            if(myReader.GetInt32(4) != 0)                                
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                        data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                        for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
                        }
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
        }

        public int countSemilla(int semana, string orden)
        {
            string query = "SELECT COUNT(ID) FROM formatoSemilla WHERE Semana = " + semana + " AND Orden = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            int count = 0;
            try
            {
                while (myReader.Read())
                {
                    count = myReader.GetInt32(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return count;
        }

        public void cargarADFExtraccion(int semana, string orden, string adf, DataGridView data)
        {
            string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "'";
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
                    data.Rows.Add();
                    if (!myReader.IsDBNull(4))
                        if (myReader.GetInt32(4) != 0)
                            data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                    data.Rows[i].Cells[1].Value = myReader.GetInt32(0);
                    data.Rows[i].Cells[0].Value = i + 1;
                    data.Rows[i].Cells[2].Value = myReader.GetString(5);
                    for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                    {
                        data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
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
        }

        public void cargarSemilla(int semana, string orden, DataGridView data)
        {
            string query = "SELECT * FROM formatoSemilla WHERE Semana = " + semana + " AND Orden = " + orden;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            semillas = 0;
            try
            {
                while (myReader.Read())
                {
                    empleadosSemillas[semillas] = myReader.GetInt32(3);
                    especiesSemillas[semillas] = myReader.GetInt32(4);
                    data.Rows.Add();
                    data.Rows[semillas].Cells[0].Value = semillas+1;
                    data.Rows[semillas].Cells[1].Value = myReader.GetInt32(0);
                    for (int j = 5, k = 5; j < 11; j++, k++)
                    { 
                        data.Rows[semillas].Cells[k].Value = myReader.GetInt32(j).ToString();
                    }
                    semillas++;
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

        public void cargarADFEquipos(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count-1; i++)
            {
                string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'";
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
                        if (!myReader.IsDBNull(4))
                            if (myReader.GetInt32(4) != 0)
                                data.Rows[i].Cells[3].Value = myReader.GetInt32(4);
                        for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
                        }
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
        }

        public void cargarADF2(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                string query = "SELECT * FROM Control WHERE Semana = " + semana + " AND Orden = " + orden + " AND Detalle = '" + data.Rows[i].Cells[2].Value.ToString() + "' AND ADF = '" + adf + "'";
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
                        if (!myReader.IsDBNull(4))
                            data.Rows[i].Cells[3].Value = myReader.GetString(4);
                        for (int j = 6, k = 5; j < 12; j++, k = k + 2)
                        {
                            data.Rows[i].Cells[k].Value = myReader.GetInt32(j).ToString();
                        }
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
        }

        public void crearSemilla(int semana, string orden, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count-1; i++)
            {
                conn.ConnectionString = connectionString;
                OleDbCommand cmd = new OleDbCommand("INSERT INTO formatoSemilla(Semana,Orden,Trabajador,Semilla,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable) VALUES (@Semana,@Orden,@Trabajador,@Semilla,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable)");
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Trabajador", OleDbType.VarChar).Value = "";
                    cmd.Parameters.Add("@Semilla", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[6].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[8].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[10].Value.ToString();
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
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

        public void crearADF(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                int j = 0;
                OleDbCommand cmd;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 1;
                }
                if (adf.Contains("ADF006") && !adf.Equals("ADF006-2"))
                {
                    if (i == 0)
                        adf = "ADF006-1";
                    else
                        adf = "ADF006";
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if(j==0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();                    
                    cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
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

        public void crearADF(int semana, string orden, string adf, DataGridView data, int maquina, ComboBox empleado, DateTimePicker date, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (data.Rows[i].Cells[2].Value != null)
                {
                    conn.ConnectionString = connectionString;
                    int j = 0;
                    OleDbCommand cmd;
                    cmd = new OleDbCommand("INSERT INTO formatoMantenimiento(Semana,Orden,Fecha,Equipo,Detalle,Insumo,Cantidad,Tipo,Estado,Editable,ADF,Empleado) VALUES (@Semana,@Orden,@Fecha,@Equipo,@Detalle,@Insumo,@Cantidad,@Tipo,@Estado,@Editable,@ADF,@Empleado)");
                    cmd.Connection = conn;
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                        cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                        cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = date.Value.ToString("dd") + "/" + date.Value.ToString("MM") + "/" + date.Value.Year;
                        cmd.Parameters.Add("@Equipo", OleDbType.VarChar).Value = maquina;
                        if (data.Rows[i].Cells[1].Value != null)
                            cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value.ToString();
                        else
                            cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                        if (data.Rows[i].Cells[2].Value != null)
                            cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                        cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = tipo;
                        cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
                        cmd.Parameters.Add("@Empleado", OleDbType.VarChar).Value = empleado.SelectedValue;
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
        }

        public void modificarADF(int semana, string orden, string adf, DataGridView data, int maquina, ComboBox empleado, DateTimePicker date, string tipo)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                if (data.Rows[i].Cells[2].Value != null)
                {
                    conn.ConnectionString = connectionString;
                    int j = 0;
                    OleDbCommand cmd;
                    if (data.Rows[i].Cells[0].Value != null)
                    {
                        cmd = new OleDbCommand("UPDATE formatoMantenimiento SET Fecha=@Fecha,Equipo=@Equipo,Detalle=@Detalle,Insumo=@Insumo,Cantidad=@Cantidad,Tipo=@Tipo,Estado=@Estado,Editable=@Editable,ADF=@ADF,Empleado=@Empleado WHERE ID = " + data.Rows[i].Cells[0].Value);
                    }
                    else
                    {
                        cmd = new OleDbCommand("INSERT INTO formatoMantenimiento(Semana,Orden,Fecha,Equipo,Detalle,Insumo,Cantidad,Tipo,Estado,Editable,ADF,Empleado) VALUES (@Semana,@Orden,@Fecha,@Equipo,@Detalle,@Insumo,@Cantidad,@Tipo,@Estado,@Editable,@ADF,@Empleado)");
                        j = 1;
                    }                    
                    cmd.Connection = conn;
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        if (j == 1)
                        {
                            cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                            cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                        }
                        cmd.Parameters.Add("@Fecha", OleDbType.VarChar).Value = date.Value.ToString("dd") + "/" + date.Value.ToString("MM") + "/" + date.Value.Year;
                        cmd.Parameters.Add("@Equipo", OleDbType.VarChar).Value = maquina;
                        if (data.Rows[i].Cells[1].Value != null)
                            cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[1].Value.ToString();
                        else
                            cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                        if (data.Rows[i].Cells[2].Value != null)
                            cmd.Parameters.Add("@Insumo", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                        cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                        cmd.Parameters.Add("@Tipo", OleDbType.VarChar).Value = tipo;
                        cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                        cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
                        cmd.Parameters.Add("@Empleado", OleDbType.VarChar).Value = empleado.SelectedValue;
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
        }

        public void modificarHorometro(int maquina, int horometro)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE Maquinarias SET Horometro=@Horometro WHERE ID = " + maquina);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Horometro", OleDbType.VarChar).Value = horometro;
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

        public void crearADF2(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count-1; i++)
            {
                conn.ConnectionString = connectionString;
                int j = 0;
                OleDbCommand cmd;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    if (data.Rows[i].Cells[5].Value != null)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[7].Value != null)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[9].Value != null)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[11].Value != null)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[13].Value != null)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[15].Value != null)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
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

        public void crearADF3(int semana, string orden, string adf, DataGridView data)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                conn.ConnectionString = connectionString;
                int j = 0;
                OleDbCommand cmd;
                if (data.Rows[i].Cells[3].Value != null)
                {
                    cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Modelo,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Modelo,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 0;
                }
                else
                {
                    cmd = new OleDbCommand("INSERT INTO Control(Semana,Unidad,Orden,Detalle,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,Estado,Editable,ADF) VALUES (@Semana,@Unidad,@Orden,@Detalle,@Lunes,@Martes,@Miercoles,@Jueves,@Viernes,@Sabado,@Estado,@Editable,@ADF)");
                    j = 1;
                }
                cmd.Connection = conn;
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Unidad", OleDbType.VarChar).Value = data.Rows[i].Cells[4].Value.ToString();
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    if (data.Rows[i].Cells[2].Value != null)
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = data.Rows[i].Cells[2].Value.ToString();
                    else
                        cmd.Parameters.Add("@Detalle", OleDbType.VarChar).Value = "";
                    if (j == 0)
                        cmd.Parameters.Add("@Modelo", OleDbType.VarChar).Value = data.Rows[i].Cells[3].Value.ToString();
                    if (data.Rows[i].Cells[5].Value != null)
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = data.Rows[i].Cells[5].Value.ToString();
                    else
                        cmd.Parameters.Add("@Lunes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[7].Value != null)
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = data.Rows[i].Cells[7].Value.ToString();
                    else
                        cmd.Parameters.Add("@Martes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[9].Value != null)
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = data.Rows[i].Cells[9].Value.ToString();
                    else
                        cmd.Parameters.Add("@Miercoles", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[11].Value != null)
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = data.Rows[i].Cells[11].Value.ToString();
                    else
                        cmd.Parameters.Add("@Jueves", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[13].Value != null)
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = data.Rows[i].Cells[13].Value.ToString();
                    else
                        cmd.Parameters.Add("@Viernes", OleDbType.VarChar).Value = 0;
                    if (data.Rows[i].Cells[15].Value != null)
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = data.Rows[i].Cells[15].Value.ToString();
                    else
                        cmd.Parameters.Add("@Sabado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Estado", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Editable", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
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

        public void formato(DataGridView data, string[] formato)
        {
            for (int i = 0; i < formato.Length; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = i + 1;
                data.Rows[i].Cells[1].Value = formato[i];
            }
        }

        public void formatoMant(DataGridView data, string[] formato)
        {
            for (int i = 0; i < formato.Length; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[1].Value = formato[i];
            }
        }

        public void formato(DataGridView data, int inicio)
        {
            for (int i = 0; i < 75; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = inicio + i + 1;
                data.Rows[i].Cells[1].Value = "0";
                data.Rows[i].Cells[2].Value = "0";
                data.Rows[i].Cells[3].Value = "0";
            }
        }

        public void formato(DataGridView data, string formato)
        {
            data.Rows.Add();
            for (int j = 4; j < 16; j = j + 2)
            {
                data.Rows[0].Cells[j].Value = formato;
                data.Rows[0].Cells[j + 1].Value = "0";
            }
        }

        public void formato(DataGridView data, string formato, int column)
        {
            for (int j = 4; j < 16; j = j + 2)
            {
                data.Rows[column].Cells[j].Value = formato;
            }
        }

        public void formato(DataGridView data, string[,] formato)
        {
            for (int i = 0; i < (formato.Length) / 2; i++)
            {
                data.Rows.Add();
                data.Rows[i].Cells[0].Value = i + 1;
                data.Rows[i].Cells[2].Value = formato[0, i];
                for (int j = 4; j < 16; j = j + 2)
                {
                    data.Rows[i].Cells[j].Value = formato[1, i];
                }
                for (int j = 5; j < 17; j = j + 2)
                {
                    data.Rows[i].Cells[j].Value = "0";
                }
            }
        }

        public void getEmpleados(string orden)
        {

            string query = "SELECT t.ID, (t.Nombres + ' ' + t.Apellidos) FROM Trabajadores AS t INNER JOIN ordenEmpleados AS o ON t.ID = o.Trabajador WHERE o.Orden = " + orden;
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
                    area = myReader.GetInt32(0).ToString();
                    estado = myReader.GetString(1);
                    supervisor = myReader.GetString(2);
                    predio = myReader.GetString(3);
                    lote = myReader.GetString(4);
                    actividad = myReader.GetString(5);
                    unidad = myReader.GetString(6);
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

        public void getOrden(string orden)
        {
            string query = "SELECT h.Area, h.estadoLote, (t.Nombres+' '+t.Apellidos), b.Predio, Areas.Lote, a.Actividad, a.Unidad_de_Medida,h.fechaInicio,h.fechaFinal,Areas.Codigo,a.Descripcion_Actividad,h.Cliente,h.Transportador, h.Operador, h.OT FROM ((historicoOrdenes AS h INNER JOIN Trabajadores AS t ON h.Supervisor = t.ID) INNER JOIN Actividades AS a ON h.Actividad = a.ID) INNER JOIN (BancoTierras AS b INNER JOIN Areas ON b.ID = Areas.Predio) ON h.Lote = Areas.Codigo WHERE h.ID = " + orden + " UNION ALL SELECT h.Area, h.estadoLote, (t.Nombres+' '+t.Apellidos), b.Predio, Lotes.Lote, a.Actividad, a.Unidad_de_Medida,h.fechaInicio,h.fechaFinal, Lotes.Codigo,a.Descripcion_Actividad,h.Cliente,h.Transportador, h.Operador, h.OT FROM ((historicoOrdenes AS h INNER JOIN Trabajadores AS t ON h.Supervisor = t.ID) INNER JOIN Actividades AS a ON h.Actividad = a.ID) INNER JOIN (BancoTierras AS b INNER JOIN Lotes ON b.ID = Lotes.Predio) ON h.Lote = Lotes.Codigo WHERE h.ID = " + orden + " UNION ALL SELECT h.Area, h.estadoLote, (t.Nombres+' '+t.Apellidos), b.Predio, LoteGanadero.Lote, a.Actividad, a.Unidad_de_Medida,h.fechaInicio,h.fechaFinal, LoteGanadero.Codigo,a.Descripcion_Actividad,h.Cliente,h.Transportador, h.Operador, h.OT FROM ((historicoOrdenes AS h INNER JOIN Trabajadores AS t ON h.Supervisor = t.ID) INNER JOIN Actividades AS a ON h.Actividad = a.ID) INNER JOIN (BancoTierras AS b INNER JOIN LoteGanadero ON b.ID = LoteGanadero.Predio) ON h.Lote = LoteGanadero.Codigo WHERE h.ID = " + orden;
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
                    area = myReader.GetInt32(0).ToString();
                    estado = myReader.GetString(1);
                    supervisor = myReader.GetString(2);
                    predio = myReader.GetString(3);
                    lote = myReader.GetString(4);
                    actividad = myReader.GetString(5);
                    unidad = myReader.GetString(6);
                    fechaInicial = myReader.GetString(7);
                    fechaFinal = myReader.GetString(8);
                    codigo = myReader.GetInt32(9);
                    descripcion = myReader.GetString(10);
                    cliente = myReader.GetInt32(11);
                    transportador = myReader.GetInt32(12);
                    operador = myReader.GetInt32(13);
                    nomOrden = myReader.GetString(14);
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

        public void getLote(int codigo)
        {
            string query = "SELECT Especie,estadoFSC,FSC FROM Lotes WHERE Codigo = " + codigo;
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
                    especie = myReader.GetString(0);
                    estadoFSC = myReader.GetString(1);
                    FSC = myReader.GetString(2);
                }
                else
                {
                    especie = "";
                    estadoFSC = "";
                    FSC = "";
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

        public void getOperador(int operador)
        {
            string query = "SELECT Operador,NIT FROM Operador WHERE ID = " + operador;
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
                    nomOperador = myReader.GetString(0);
                    NIT = myReader.GetString(1);
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

        public void hideFormatos()
        {
            for (int i = 0; i < tabControl1.TabPages.Count; i++)
            {
                ((Control)this.tabControl1.TabPages[i]).Enabled = false;
            }
        }

        public void eliminarFormatos()
        {
            for (int i = 0; i < tabControl1.TabPages.Count; i++)
            {
                if (((Control)this.tabControl1.TabPages[i]).Enabled == false)
                {
                    if (tabControl1.HasChildren)
                    {
                        tabControl1.TabPages.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        sw2 = false;
                    }
                }
            }
        }

        public void getFormatos(string orden)
        {
            string query = "SELECT f.Formato, f.Actividad FROM formatosActividad AS f INNER JOIN historicoOrdenes AS h ON f.Actividad = h.Actividad WHERE h.ID = " + orden + " ORDER BY h.ID desc";
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
                    for (int i = 0; i < tabControl1.TabPages.Count; i++)
                    {
                        if (tabControl1.TabPages[i].Text.Equals(myReader.GetString(0)))
                        {
                            ((Control)this.tabControl1.TabPages[i]).Enabled = true;
                        }
                    }
                }
                if (!myReader.HasRows)
                {
                    MessageBox.Show("La actividad no contiene formatos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Close();
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

        public void formatoTipo1(Microsoft.Office.Interop.Excel.Application XcelApp, int semana)
        {
            XcelApp.Cells[5, 2] = supervisor;
            XcelApp.Cells[5, 8] = DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Year;
            XcelApp.Cells[10, 2] = actividad;
            XcelApp.Cells[13, 2] = lote;
            XcelApp.Cells[14, 2] = predio;
            XcelApp.Cells[14, 4] = area;
            XcelApp.Cells[14, 5] = unidad;
            XcelApp.Cells[14, 7] = estado;
            XcelApp.Cells[14, 12] = nomOrden;
            XcelApp.Cells[5, 12] = semana;
        }

        public void formatoTipo2(Microsoft.Office.Interop.Excel.Application XcelApp, int semana)
        {
            XcelApp.Cells[5, 2] = supervisor;
            XcelApp.Cells[5, 9] = DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Year;
            XcelApp.Cells[10, 2] = actividad;
            XcelApp.Cells[9, 2] = lote;
            XcelApp.Cells[10, 2] = predio;
            XcelApp.Cells[10, 5] = area;
            XcelApp.Cells[10, 6] = unidad;
            XcelApp.Cells[10, 8] = estado;
            XcelApp.Cells[10, 12] = nomOrden;
            XcelApp.Cells[5, 14] = semana;
        }

        public void formatoTipo5(Microsoft.Office.Interop.Excel.Application XcelApp, int semana)
        {
            XcelApp.Cells[5, 2] = supervisor;
            XcelApp.Cells[5, 9] = DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Year;
            XcelApp.Cells[11, 2] = lote;
            XcelApp.Cells[12, 2] = predio;
            XcelApp.Cells[12, 5] = area;
            XcelApp.Cells[12, 6] = unidad;
            XcelApp.Cells[12, 8] = estado;
            XcelApp.Cells[12, 12] = nomOrden;
            XcelApp.Cells[5, 14] = semana;
        }

        public void formatoTipo3(Microsoft.Office.Interop.Excel.Application XcelApp, int semana)
        {
            XcelApp.Cells[5, 2] = supervisor;
            XcelApp.Cells[5, 9] = DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Year;
            XcelApp.Cells[10, 2] = actividad;
            XcelApp.Cells[10, 2] = lote;
            XcelApp.Cells[10, 5] = area;
            XcelApp.Cells[10, 6] = unidad;
            XcelApp.Cells[10, 8] = estado;
            XcelApp.Cells[10, 12] = nomOrden;
            XcelApp.Cells[5, 12] = semana;
        }

        public void formatoTipo4(Microsoft.Office.Interop.Excel.Application XcelApp, int semana)
        {
            XcelApp.Cells[6, 2] = DateTime.Now.Day + " / " + DateTime.Now.Month + " / " + DateTime.Now.Year;
            XcelApp.Cells[6, 7] = nomOrden;
            XcelApp.Cells[7, 2] = codigo;
            XcelApp.Cells[7, 7] = lote;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmCrearOrden newFrm = new frmCrearOrden(OT,tipousuario);
            if (!newFrm.IsDisposed)
            {
                this.Hide();
                newFrm.ShowDialog();
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = DateTime.Now;
            Calendar cal = dfi.Calendar;
            int semana = cal.GetWeekOfYear(date1, dfi.CalendarWeekRule, dfi.FirstDayOfWeek) - 1;
            Directory.CreateDirectory("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos");
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();

            if (tabControl1.SelectedTab.Text.Equals("ADF002"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF002*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;                
                formatoTipo1(XcelApp, semana);
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    for (int j = 0; j < 3; j++)
                        XcelApp.Cells[18 + i, 2 + j] = dataGridView3.Rows[i].Cells[2 + j].Value;   
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF003"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF003*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo2(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF004"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF004*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo2(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF005"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF005*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo5(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF006"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF006*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo5(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF007"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF007*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo2(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF008"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF008*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo2(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF009"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF009*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo2(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF010"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF010*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo3(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF011"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF011*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo2(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF012"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF012*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo4(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF013"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF013*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo4(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF014"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF014*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo4(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF015"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF015*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo4(XcelApp, semana);
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF016"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF016*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                XcelApp.Cells[5, 7] = dateTimePicker1.Text;

            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF017"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF017*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                XcelApp.Cells[5, 7] = dateTimePicker2.Text;
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF018"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF018*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                XcelApp.Cells[5, 7] = dateTimePicker3.Text;
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF019"))
            {
                string[] prueba = Directory.GetFiles("C:\\Users\\" + Environment.UserName + "\\Dropbox\\Formatos\\", "ADF019*");
                XcelApp.Application.Workbooks.Add(prueba[0]);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)XcelApp.Worksheets.get_Item(1);
                xlWorkSheet.PageSetup.LeftHeader = "&B" + nomOperador + "\n" + NIT;
                formatoTipo4(XcelApp, semana);
                XcelApp.Cells[11, 3] = actividad;
                XcelApp.Cells[12, 3] = descripcion;
            }
            XcelApp.Visible = true;
        }

        public bool tabExiste(string tab)
        {
            bool sw = false;
            for (int i = 0; i < tabControl1.TabPages.Count; i++)
            {
                if (tabControl1.TabPages[i].Text.Equals(tab))
                    sw = true;
            }
            return sw;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Text.Equals("ADF002"))
            {

            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF013"))
            {
                if (tabExiste("ADF006"))
                {
                    if (dataGridView9.Rows[0].Cells[3].Value != null)
                    {
                        comboBox3.SelectedValue = dataGridView9.Rows[0].Cells[3].Value;
                        treeView1.Nodes.Clear();
                        llenarArbol(treeView1);
                        cargarADF(semanaFinal, OT, "ADF006-2", treeView1);
                    }
                }
                else if (tabExiste("ADF005"))
                {
                    if (dataGridView4.Rows[0].Cells[3].Value != null)
                    {
                        comboBox3.SelectedValue = dataGridView4.Rows[0].Cells[3].Value;
                        treeView1.Nodes.Clear();
                        llenarArbol(treeView1);
                        cargarADF(semanaFinal, OT, "ADF005-2", treeView1);
                    }
                }
            }
            else if (tabControl1.SelectedTab.Text.Equals("ADF015"))
            {
                if (tabExiste("ADF006"))
                {
                    if (dataGridView9.Rows[0].Cells[3].Value != null)
                    {
                        //comboBox10.SelectedValue = dataGridView9.Rows[0].Cells[3].Value;
                        treeView2.Nodes.Clear();
                        llenarArbol(treeView2);
                        cargarADF(semanaFinal, OT, "ADF006-2", treeView2);
                    }
                }
                else if (tabExiste("ADF005"))
                {
                    if (dataGridView4.Rows[0].Cells[3].Value != null)
                    {
                        //comboBox10.SelectedValue = dataGridView4.Rows[0].Cells[3].Value;
                        treeView2.Nodes.Clear();
                        llenarArbol(treeView2);
                        cargarADF(semanaFinal, OT, "ADF005-2", treeView2);
                    }
                }
                
            }
        }

        public void llenarArbol(TreeView tree) 
        {
            tree.Nodes.Add("Lunes");
            tree.Nodes.Add("Martes");
            tree.Nodes.Add("Miercoles");
            tree.Nodes.Add("Jueves");
            tree.Nodes.Add("Viernes");
            tree.Nodes.Add("Sabado");
        }

        public void selectAll(DataGridView data, int inicio, int final)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                DataGridViewCheckBoxCell ch2 = new DataGridViewCheckBoxCell();
                ch2 = (DataGridViewCheckBoxCell)data.Rows[i].Cells[12];
                if ((bool)ch2.FormattedValue == true && sw == false)
                {
                    for (int j = inicio; j < final; j++)
                    {
                        data.Rows[i].Cells[j].Value = true;
                    }
                    data.Rows[i].Cells[data.Columns.Count - 1].Value = 6;
                    data.Rows[i].Cells[data.Columns.Count - 2].Value = false;
                    sw = true;
                }
                else
                {
                    sw = false;
                }

            }
        }

        public void Contador()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                int total = 0;
                for (int j = 5; j < 12; j++)
                {
                    DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                    ch1 = (DataGridViewCheckBoxCell)dataGridView3.Rows[i].Cells[j];
                    if ((bool)ch1.FormattedValue == true)
                        total++;
                }
                dataGridView3.Rows[i].Cells[13].Value = total;
            }
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            selectAll(dataGridView3, 5, 11);
            Contador();
        }

        private void dataGridView3_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView3.IsCurrentCellDirty)
            {
                dataGridView3.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        public void ContadorSemanal(DataGridView data, int inicio, int final)
        {
            if (data.Rows.Count != 0)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    int total = 0;
                    for (int j = inicio; j < final; j++)
                    {
                        int num;

                        if (!(data.Rows[i].Cells[j].Value == null))
                        {
                            bool isNum = Int32.TryParse(data.Rows[i].Cells[j].Value.ToString(), out num);
                            if (isNum)
                            {
                                total = total + Int32.Parse(data.Rows[i].Cells[j].Value.ToString());
                            }
                        }
                    }
                    data.Rows[i].Cells[data.Columns.Count - 1].Value = total;
                }
            }
        }

        public void Contador(DataGridView data, int inicio, int final)
        {
            if (data.Rows.Count != 0)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    int total = 0;
                    for (int j = inicio; j < final; j = j + 2)
                    {
                        int num;

                        if (!(data.Rows[i].Cells[j].Value == null))
                        {
                            bool isNum = Int32.TryParse(data.Rows[i].Cells[j].Value.ToString(), out num);
                            if (isNum)
                            {
                                total = total + Int32.Parse(data.Rows[i].Cells[j].Value.ToString());
                            }
                        }
                    }
                    data.Rows[i].Cells[data.Columns.Count - 1].Value = total;
                }
            }
        }

        public void Contador(DataGridView data, int inicio, int final, string[,] adf)
        {
            if (data.Rows.Count == (adf.Length) / 2)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    int total = 0;
                    for (int j = inicio; j < final; j = j + 2)
                    {
                        int num;
                        if (!(data.Rows[i].Cells[j].Value == null))
                        {
                            bool isNum = Int32.TryParse(data.Rows[i].Cells[j].Value.ToString(), out num);
                            if (isNum)
                            {
                                total = total + Int32.Parse(data.Rows[i].Cells[j].Value.ToString());
                            }
                        }
                    }
                    data.Rows[i].Cells[data.Columns.Count - 1].Value = total;
                }
            }
        }

        private void dataGridView4_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView4, 5, 17, ADF005);
        }

        private void dataGridView4_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView4);
        }

        public void dirtyCell(DataGridView data)
        {
            if (data.IsCurrentCellDirty)
                data.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF004") == false)
            {
                crearADF(semanaFinal, OT, "ADF004", dataGridView2);
                //eliminarADF(semanaFinal, OT, "ADF004");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF004"), "+");            
                agregarLog("Usuario ha modificado el formato ADF004 en la OT #" + OT, userName);
            }
            else
            {
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF004"), "-");            
                //agregarADF(semanaFinal, OT, "ADF004");
                modificarADF(semanaFinal, OT, "ADF004", dataGridView2);
                //eliminarADF(semanaFinal, OT, "ADF004");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF004"), "+");            
                agregarLog("Usuario ha modificado el formato ADF004 en la OT #" + OT, userName);
            }
            //if (sw2)
            //{
            //    agregarADF(modificarADF002(semana, DateTime.Now.Year, supervisor);
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            MessageBox.Show("Control de Insumos Extracción Mecánica registrado.");
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView2, 5, 17, ADF004);
        }

        private void dataGridView4_Enter(object sender, EventArgs e)
        {
        }

        private void dataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView2);
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView1);
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView1, 5, 17, ADF003);
        }

        private void dataGridView9_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView9, 5, 17, ADF006);
        }

        private void dataGridView9_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView9);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ADF002Existe(semanaFinal, OT) == false)
            {
                crearADF002(semanaFinal, OT);
                agregarDias(semanaFinal, OT);
                agregarCostoJornal(Int32.Parse(OT), dataGridView3);
                agregarLog("Usuario ha modificado el formato ADF002 en la OT #" + OT, userName);
            }
            else
            {
                eliminarCostoJornal(semanaFinal, OT);
                eliminarDias(semanaFinal, OT);
                modificarADF002(semanaFinal, OT);
                agregarDias(semanaFinal, OT);
                agregarCostoJornal(Int32.Parse(OT), dataGridView3);
                agregarLog("Usuario ha modificado el formato ADF002 en la OT #" + OT, userName);
            }
            //if (sw2)
            //{
            //    agregarADF(modificarADF002(semana, DateTime.Now.Year, supervisor);
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            MessageBox.Show("Asistencia registrada.");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF003") == false)
            {
                crearADF(semanaFinal, OT, "ADF003", dataGridView1);
                //eliminarADF(semanaFinal, OT, "ADF003");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF003"), "+");            
                agregarLog("Usuario ha modificado el formato ADF003 en la OT #" + OT, userName);
            }
            else
            {
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF003"), "-");            
                //agregarADF(semanaFinal, OT, "ADF003");
                modificarADF(semanaFinal, OT, "ADF003", dataGridView1);
                //eliminarADF(semanaFinal, OT, "ADF003");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF003"), "+");            
                agregarLog("Usuario ha modificado el formato ADF003 en la OT #" + OT, userName);
            }
            //if (sw2)
            //{
            //    agregarADF(modificarADF002(semana, DateTime.Now.Year, supervisor);
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            MessageBox.Show("Control de Insumos Extracción Manual registrado.");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF005") == false)
            {
                crearADF(semanaFinal, OT, "ADF005", dataGridView4);
                crearADF(semanaFinal, OT, "ADF005-2", dataGridView5);
                crearDaños(semanaFinal, OT, "ADF005", dataGridView6);
                //eliminarADF(semanaFinal, OT, "ADF005");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF005"), "+");            
                agregarLog("Usuario ha modificado el formato ADF005 en la OT #" + OT, userName);
            }
            else
            {
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF005"), "-");            
                //agregarADF(semanaFinal, OT, "ADF005");
                modificarADF(semanaFinal, OT, "ADF005", dataGridView4);
                modificarADF2(semanaFinal, OT, "ADF005-2", dataGridView5);
                modificarDaños(semanaFinal, OT, "ADF005", dataGridView6);
                //eliminarADF(semanaFinal, OT, "ADF005");                
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF005"), "+");            
                agregarLog("Usuario ha modificado el formato ADF005 en la OT #" + OT, userName);
            }
            //if (sw2)
            //{
            //    agregarADF(modificarADF002(semana, DateTime.Now.Year, supervisor);
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            MessageBox.Show("Control de Cosechador registrado.");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF006") == false)
            {
                crearADF(semanaFinal, OT, "ADF006", dataGridView9);
                crearADF(semanaFinal, OT, "ADF006-2", dataGridView8);
                crearDaños(semanaFinal, OT, "ADF006", dataGridView7);
                agregarADFEquipo(semanaFinal, OT, "ADF006-1");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF006"), "+");            
                //eliminarADF(semanaFinal, OT, "ADF006");
                agregarLog("Usuario ha modificado el formato ADF006 en la OT #" + OT, userName);
            }
            else
            {
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF006"), "-");
                eliminarADFEquipo(semanaFinal, OT, "ADF006-1");
                //agregarADF(semanaFinal, OT, "ADF006");
                modificarADF(semanaFinal, OT, "ADF006", dataGridView9);
                modificarADF2(semanaFinal, OT, "ADF006-2", dataGridView8);
                modificarDaños(semanaFinal, OT, "ADF006", dataGridView7);
                agregarADFEquipo(semanaFinal, OT, "ADF006-1");
                //eliminarADF(semanaFinal, OT, "ADF006");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF006"), "+");            
                agregarLog("Usuario ha modificado el formato ADF006 en la OT #" + OT, userName);
            }
            //if (sw2)
            //{
            //    agregarADF(modificarADF002(semana, DateTime.Now.Year, supervisor);
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            if ((int)(dataGridView8.Rows[0].Cells[dataGridView8.Columns.Count - 1].Value) % 10 == 0)
            {
                MessageBox.Show("Favor llenar el ADF012.");
            }
            MessageBox.Show("Control de Equipo registrado.");
        }

        private void dataGridView8_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView8, 5, 17);
        }

        private void dataGridView8_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView8);
        }

        private void dataGridView5_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView5, 4, 16);
        }

        private void dataGridView5_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView5);
        }

        public double crearPromedio(DataGridView data, double castigo)
        {
            bool termino = false;
            double promedio = 0;
            int cantidad = 0;
            while (!termino || cantidad >= data.Rows.Count)
            {
                double d1 = double.Parse(data.Rows[cantidad].Cells[1].Value.ToString());
                double d2 = double.Parse(data.Rows[cantidad].Cells[2].Value.ToString());
                double l = double.Parse(data.Rows[cantidad].Cells[3].Value.ToString());                
                if(d1 != 0 && d2 != 0 && l != 0)
                {
                    cantidad++;
                    double promedioTemp = ((((((d1 + d2 - (2 * (castigo)))/200))* ((d1 + d2 - (2 * (castigo)))/200))) * Math.PI) * l;
                    promedio += promedioTemp;
                }
                else
                {
                    termino = true;
                }
            }
            if (cantidad != 0)
                promedio = (promedio / cantidad);
            else
                promedio = 0; 
            return promedio;
        }

        public double PromedioDiametro(DataGridView data)
        {
            bool termino = false;
            double promedio = 0;
            int cantidad = 0;
            while (!termino || cantidad >= data.Rows.Count)
            {
                double d1 = double.Parse(data.Rows[cantidad].Cells[1].Value.ToString());
                double d2 = double.Parse(data.Rows[cantidad].Cells[2].Value.ToString());
                double l = double.Parse(data.Rows[cantidad].Cells[3].Value.ToString());                
                if (d1 != 0 && d2 != 0 && l != 0)
                {
                    cantidad++;
                    double promedioTemp = (d1 + d2) / 2;
                    promedio += promedioTemp;
                }
                else
                {
                    termino = true;
                }
            }
            if (cantidad != 0)
                promedio = (promedio / cantidad);
            else
                promedio = 0;
            return promedio;
        }

        public double PromedioLargo(DataGridView data)
        {
            bool termino = false;
            double promedio = 0;
            int cantidad = 0;
            while (!termino || cantidad >= data.Rows.Count)
            {
                double d1 = double.Parse(data.Rows[cantidad].Cells[1].Value.ToString());
                double d2 = double.Parse(data.Rows[cantidad].Cells[2].Value.ToString());
                double l = double.Parse(data.Rows[cantidad].Cells[3].Value.ToString());
                if (d1 != 0 && d2 != 0 && l != 0)
                {
                    cantidad++;
                    promedio += l;
                }
                else
                {
                    termino = true;
                }
            }
            if (cantidad != 0)
                promedio = (promedio / cantidad);
            else
                promedio = 0;            
            return promedio;
        }

        public int PromedioCantidad(DataGridView data)
        {
            bool termino = false;
            int cantidad = 0;
            while (!termino || cantidad >= data.Rows.Count)
            {
                double d1 = double.Parse(data.Rows[cantidad].Cells[1].Value.ToString());
                double d2 = double.Parse(data.Rows[cantidad].Cells[2].Value.ToString());
                double l = double.Parse(data.Rows[cantidad].Cells[3].Value.ToString());
                if (d1 != 0 && d2 != 0 && l != 0)
                {
                    cantidad++;
                }
                else
                {
                    termino = true;
                }
            }
            return cantidad;
        }        

        public void crearFormatoMaterial(double promedio,string tipo, string especie, double diametro, double largo, int cantidad)
        {

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

        private void button13_Click(object sender, EventArgs e)
        {
            int cant = 0;
            if (PromedioCantidad(material1) != 0)
                cant++;
            if (PromedioCantidad(material2) != 0)
                cant++;
            if (PromedioCantidad(material3) != 0)
                cant++;
            if (PromedioCantidad(material4) != 0)
                cant++;

            double promedio = (crearPromedio(material1,double.Parse(textBox15.Text)) + crearPromedio(material2,double.Parse(textBox15.Text)) + crearPromedio(material3,double.Parse(textBox15.Text)) + crearPromedio(material4,double.Parse(textBox15.Text)))/cant;
            double promedioD = (PromedioDiametro(material1) + PromedioDiametro(material2) + PromedioDiametro(material3) + PromedioDiametro(material4))/cant;
            double promedioL = (PromedioLargo(material1) + PromedioLargo(material2) + PromedioLargo(material3) + PromedioLargo(material4)) / cant;
            int cantidad = (PromedioCantidad(material1) + PromedioCantidad(material2) + PromedioCantidad(material3) + PromedioCantidad(material4));
            string raleo = "", especie = "", trailer = "";
            raleo = getRaleo(radioButton6, radioButton5, radioButton7);
            especie = getEspecie(radioButton4, radioButton3, radioButton31, textBox2);
            if (radioButton33.Checked)
                trailer = "Farmi Primero 9000";
            else
                trailer = "Pfanzelt 15100";
            agregarVolumen(promedioD, promedioL, cantidad, promedio, especie, trailer, raleo, "ADF012");
            MessageBox.Show("Transferencia de Material a Aserradero FCTH Registrado.");
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked)
            {
                label76.Visible = true;
                textBox6.Visible = true;
            }
            else
            {
                label76.Visible = false;
                textBox6.Visible = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                label67.Visible = true;
                textBox1.Visible = true;
            }
            else
            {
                label67.Visible = false;
                textBox1.Visible = false;
            }
        }

        private void dataGridView12_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Contador(dataGridView12, 5, 17, ADF009);
        }

        private void dataGridView12_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView12);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF007") == false)
            {
                crearADF2(semanaFinal, OT, "ADF007", dataGridView10);
                agregarADFEquipo(semanaFinal, OT, "ADF007");
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF007"), "-");            
                agregarLog("Usuario ha modificado el formato ADF007 en la OT #" + OT, userName);
            }
            else
            {
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF007"), "-");            
                eliminarADFEquipo(semanaFinal, OT, "ADF007");
                modificarADFEquipos(semanaFinal, OT, "ADF007", dataGridView10);
                agregarADFEquipo(semanaFinal, OT, "ADF007");
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF007"), "+");            
                agregarLog("Usuario ha modificado el formato ADF007 en la OT #" + OT, userName);
            }
            MessageBox.Show("Producción Diaria - Extracción Manual registrada.");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF009") == false)
            {
                crearADF(semanaFinal, OT, "ADF009", dataGridView12);
                //agregarADF(semanaFinal, OT, "ADF009");
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF009"), "-");            
                agregarLog("Usuario ha modificado el formato ADF009 en la OT #" + OT, userName);
            }
            else
            {
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF009"), "-");            
                //eliminarADF(semanaFinal, OT, "ADF009");
                modificarADF(semanaFinal, OT, "ADF009", dataGridView12);
                //agregarADF(semanaFinal, OT, "ADF009");
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF009"), "+");            
                agregarLog("Usuario ha modificado el formato ADF009 en la OT #" + OT, userName);
            }
            //if (sw2)
            //{
            //    agregarADF(modificarADF002(semana, DateTime.Now.Year, supervisor);
            //    agregarLog("Permiso de modificación de asistencia revocado.", usuario);
            //}
            MessageBox.Show("Control de Insumos Recoleccion de Semilla registrado.");
        }

        public void agregarDetalle(TextBox text, DataGridView data)
        {
            if (text.Text.Equals(""))
            {
                MessageBox.Show("Favor ingresar un detalle para poder registrar el mantenimiento.");
            }
            else
            {
                data.Rows.Add();
                data.Rows[data.Rows.Count - 1].Cells[0].Value = data.Rows.Count;
                data.Rows[data.Rows.Count - 1].Cells[1].Value = text.Text;
            }
        }

        public string getCedula(int id)
        {
            string query = "SELECT Cedula FROM Trabajadores WHERE id = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string cedula = "";
            try
            {
                if (myReader.Read())
                {
                    cedula = myReader.GetInt32(0).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return cedula;
        }

        public string getCedulaTransportador(int id)
        {
            string query = "SELECT Cedula FROM Transportadores WHERE id = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string cedula = "";
            try
            {
                if (myReader.Read())
                {
                    cedula = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return cedula;
        }

        public string getPlacaTransportador(int id)
        {
            string query = "SELECT Placa FROM Transportadores WHERE id = " + id;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            OleDbDataReader myReader = cmd.ExecuteReader();
            string cedula = "";
            try
            {
                if (myReader.Read())
                {
                    cedula = myReader.GetString(0);
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return cedula;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            //agregarDetalle(textBox15, dataGridView18);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            //agregarDetalle(textBox16, dataGridView19);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            //agregarDetalle(textBox19, dataGridView23);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && !comboBox1.SelectedValue.ToString().Equals("System.Data.DataRowView"))
                textBox3.Text = getCedula(Int32.Parse(comboBox1.SelectedValue.ToString()));
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedItem != null && !comboBox4.SelectedValue.ToString().Equals("System.Data.DataRowView"))
                textBox4.Text = getCedula(Int32.Parse(comboBox4.SelectedValue.ToString()));
        }

        private void button21_Click(object sender, EventArgs e)
        {
            //if (button21.Text.Equals("Semana Anterior"))
            //{
            //    if (esEditable(OT, semanaFinal - 1) || DateTime.Now.DayOfWeek.Equals("Monday"))
            //    {
            //        frmOrdenFormatos newFrm = new frmOrdenFormatos(OT, 1);
            //        this.Hide();
            //        newFrm.ShowDialog();
            //        this.Close();
            //    }
            //    else
            //    {
            //        MessageBox.Show("No se puede editar la semana anterior, favor contactar al administrador.");
            //    }
            //}
            //else
            //{
            //    frmOrdenFormatos newFrm = new frmOrdenFormatos(OT, 0);
            //    this.Hide();
            //    newFrm.ShowDialog();
            //    this.Close();
            //}
            frmResumenOrden newFrm = new frmResumenOrden(Int32.Parse(OT));
            newFrm.Show();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            textBox5.Text = "";
        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            textBox5.Text = "";
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem != null && !comboBox3.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                string idEmpleado = empleadoMaquina(comboBox3.SelectedValue.ToString());
                if (!idEmpleado.Equals(""))
                {
                    comboBox4.SelectedValue = idEmpleado;
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem != null && !comboBox2.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                string idEmpleado = empleadoMaquina(comboBox2.SelectedValue.ToString());
                if (!idEmpleado.Equals(""))
                {
                    comboBox1.SelectedValue = idEmpleado;
                }
            }
        }

        public string empleadoMaquina(string maquina)
        {
            string idEmpleado = "";
            string query = "SELECT ID From Trabajadores WHERE Maquina = " + maquina;
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
                    idEmpleado = myReader.GetInt32(0).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return idEmpleado;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF008") == false)
            {                
                crearADF2(semanaFinal, OT, "ADF008", dataGridView11);
                agregarADFEquipo(semanaFinal, OT, "ADF008");
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF008"), "-");            
                agregarLog("Usuario ha modificado el formato ADF008 en la OT #" + OT, userName);
            }
            else
            {
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF008"), "-");            
                eliminarADFEquipo(semanaFinal, OT, "ADF008");
                modificarADFEquipos(semanaFinal, OT, "ADF008", dataGridView11);
                agregarADFEquipo(semanaFinal, OT, "ADF008");
                //controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF008"), "+");            
                agregarLog("Usuario ha modificado el formato ADF008 en la OT #" + OT, userName);
            }
            MessageBox.Show("Producción Diaria - Extracción Mecanica registrada.");
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (comboBox6.SelectedValue != null)
            {
                if (mantExiste(semanaFinal, OT, "ADF017") == false)
                {
                    crearADF(semanaFinal, OT, "ADF017", dataGridView22, Maquinarias[0], comboBox6, dateTimePicker2, "Motor");
                    crearADF(semanaFinal, OT, "ADF017", dataGridView21, Maquinarias[0], comboBox6, dateTimePicker2, "Hidraulico");
                    crearADF(semanaFinal, OT, "ADF017", dataGridView20, Maquinarias[0], comboBox6, dateTimePicker2, "Diferenciales");
                    crearADF(semanaFinal, OT, "ADF017", dataGridView19, Maquinarias[0], comboBox6, dateTimePicker2, "Adicional");
                    cargarADFMant(semanaFinal, OT, "ADF017", dataGridView22, "Motor");
                    cargarADFMant(semanaFinal, OT, "ADF017", dataGridView21, "Hidraulico");
                    cargarADFMant(semanaFinal, OT, "ADF017", dataGridView20, "Diferenciales");
                    cargarADFMantDetalle(semanaFinal, OT, "ADF017", dataGridView19, comboBox6, dateTimePicker2);
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "+");
                    //eliminarADFMant(semanaFinal, OT);
                    agregarLog("Usuario ha modificado el formato ADF017 en la OT #" + OT, userName);
                }
                else
                {
                    //agregarADFMant(semanaFinal, OT);
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "-");
                    modificarADF(semanaFinal, OT, "ADF017", dataGridView22, Maquinarias[0], comboBox6, dateTimePicker2, "Motor");
                    modificarADF(semanaFinal, OT, "ADF017", dataGridView21, Maquinarias[0], comboBox6, dateTimePicker2, "Hidraulico");
                    modificarADF(semanaFinal, OT, "ADF017", dataGridView20, Maquinarias[0], comboBox6, dateTimePicker2, "Diferenciales");
                    modificarADF(semanaFinal, OT, "ADF017", dataGridView19, Maquinarias[0], comboBox6, dateTimePicker2, "Adicional");
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "+");
                    //eliminarADFMant(semanaFinal, OT);
                    agregarLog("Usuario ha modificado el formato ADF017 en la OT #" + OT, userName);
                }
                modificarHorometro(Maquinarias[0], Int32.Parse(textBox18.Text));
                MessageBox.Show("Mantenimiento Cosechador John Deere registrado.");
            }                        
            else
            {
                MessageBox.Show("Favor seleccionar quien realizo el mantenimiento.");
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            int cant = 0;
            if (PromedioCantidad(material5) != 0)
                cant++;
            if (PromedioCantidad(material6) != 0)
                cant++;
            if (PromedioCantidad(material7) != 0)
                cant++;
            if (PromedioCantidad(material8) != 0)
                cant++;

            double promedio = (crearPromedio(material5, double.Parse(textBox16.Text)) + crearPromedio(material6, double.Parse(textBox16.Text)) + crearPromedio(material7, double.Parse(textBox16.Text)) + crearPromedio(material8, double.Parse(textBox16.Text))) / cant;
            double promedioD = (PromedioDiametro(material5) + PromedioDiametro(material6) + PromedioDiametro(material7) + PromedioDiametro(material8)) / cant;
            double promedioL = (PromedioLargo(material5) + PromedioLargo(material6) + PromedioLargo(material7) + PromedioLargo(material8)) / cant;
            int cantidad = (PromedioCantidad(material5) + PromedioCantidad(material6) + PromedioCantidad(material7) + PromedioCantidad(material8));
            string raleo = "", especie = "", trailer = "";
            raleo = "N/A";
            especie = getEspecie(radioButton21, radioButton20, radioButton29, textBox23);
            if (radioButton36.Checked)
                trailer = "Camion 600";
            else if (radioButton35.Checked)
                trailer = "Camion Doble Torque";
            else
                trailer = "Tractomula";
            agregarVolumen(promedioD, promedioL, cantidad, promedio, especie, trailer, raleo, "ADF014");
            MessageBox.Show("Transferencia de Material a Aserradero FCTH Registrado.");

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox8.SelectedItem != null && !comboBox8.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            //{
            //    string idEmpleado = empleadoMaquina(comboBox8.SelectedValue.ToString());
            //    if (!idEmpleado.Equals(""))
            //    {
            //        comboBox9.SelectedValue = idEmpleado;
            //    }
            //}
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox9.SelectedItem != null && !comboBox9.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                textBox22.Text = getCedulaTransportador(Int32.Parse(comboBox9.SelectedValue.ToString()));
                textBox14.Text = getPlacaTransportador(Int32.Parse(comboBox9.SelectedValue.ToString()));
            }
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox10.SelectedItem != null && !comboBox10.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            //{
            //    string idEmpleado = empleadoMaquina(comboBox10.SelectedValue.ToString());
            //    if (!idEmpleado.Equals(""))
            //    {
            //        comboBox11.SelectedValue = idEmpleado;
            //    }
            //}
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox11.SelectedItem != null && !comboBox11.SelectedValue.ToString().Equals("System.Data.DataRowView"))
            {
                textBox35.Text = getCedulaTransportador(Int32.Parse(comboBox11.SelectedValue.ToString()));
                textBox13.Text = getPlacaTransportador(Int32.Parse(comboBox11.SelectedValue.ToString()));
            }
        }

        private void dataGridView10_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView10.Columns.Count > 5)
            {
                dataGridView10.Rows[dataGridView10.Rows.Count - 1].Cells[0].Value = dataGridView10.Rows.Count;
                dataGridView10.Rows[dataGridView10.Rows.Count - 1].Cells[2].Value = "Equipo: ";
            }
        }

        private void dataGridView10_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView10.Columns.Count > 10)
            {
                for (int i = 0; i < dataGridView10.Rows.Count; i++)
                {
                    if (dataGridView10.Rows[i].Cells[3].FormattedValue != null)
                    {
                        if (dataGridView10.Rows[i].Cells[3].FormattedValue.ToString().Contains("Motosierra") || dataGridView10.Rows[i].Cells[3].FormattedValue.ToString().Contains("Cosechador"))
                            formato(dataGridView10, "Arboles", i);
                        else
                            formato(dataGridView10, "Recorridos", i);
                    }
                }
            }
            Contador(dataGridView10, 5, 17);
        }

        private void dataGridView10_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView10);
        }

        private void dataGridView11_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView11.Columns.Count > 10)
            {
                for (int i = 0; i < dataGridView11.Rows.Count; i++)
                {
                    if (dataGridView11.Rows[i].Cells[3].FormattedValue != null)
                    {
                        if (dataGridView11.Rows[i].Cells[3].FormattedValue.ToString().Contains("Motosierra"))
                            formato(dataGridView11, "Arboles", i);
                        else
                            formato(dataGridView11, "Recorridos", i);
                    }
                }
                Contador(dataGridView11, 5, 17);
            }
        }

        private void dataGridView11_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView11);
        }

        private void dataGridView11_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView11.Columns.Count > 5)
            {
                dataGridView11.Rows[dataGridView10.Rows.Count - 1].Cells[0].Value = dataGridView11.Rows.Count;
                dataGridView11.Rows[dataGridView10.Rows.Count - 1].Cells[2].Value = "Equipo: ";
            }
        }

        private void radioButton19_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton19.Checked)
            {
                label120.Visible = true;
                textBox37.Visible = true;
            }
            else
            {
                label120.Visible = false;
                textBox37.Visible = false;
            }
        }

        private void radioButton28_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton28.Checked)
            {
                textBox36.Visible = true;
            }
            else
            {
                textBox36.Visible = false;
            }
        }

        private void radioButton31_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton31.Checked)
            {
                textBox2.Visible = true;
            }
            else
            {
                textBox2.Visible = false;
            }
        }

        private void radioButton29_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton29.Checked)
            {
                textBox23.Visible = true;
            }
            else
            {
                textBox23.Visible = false;
            }
        }

        private void radioButton30_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton30.Checked)
            {
                textBox5.Visible = true;
            }
            else
            {
                textBox5.Visible = false;
            }
        }

        private void radioButton22_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton22.Checked)
            {
                label105.Visible = true;
                textBox24.Visible = true;
            }
            else
            {
                label105.Visible = false;
                textBox24.Visible = false;
            }
        }

        private void label105_Click(object sender, EventArgs e)
        {

        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView1.SelectedNode.Parent != null)
            {
                label68.Text = "Dia: " + treeView1.SelectedNode.Parent.Text + " Recorrido #: " + treeView1.SelectedNode.Text;
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                getOrden(semanaFinal, OT, "ADF013", treeView1);
            }
        }

        public void getOrden(int semana, string orden, string adf, TreeView tree)
        {
            string query = "SELECT * FROM Transferencia WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "' AND Dia = '" + tree.SelectedNode.Parent.Text + "' AND Recorrido = " + tree.SelectedNode.Text;
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
                    if (myReader.GetInt32(5) == 0)                     
                        radioButton14.Checked = true;                                            
                    else                    
                        radioButton13.Checked = true;                    
                    textBox6.Text = myReader.GetInt32(5).ToString();
                    if (myReader.GetString(6).Equals("Teca")) 
                        radioButton11.Checked = true;                    
                    else if (myReader.GetString(6).Equals("Melina"))                    
                        radioButton12.Checked = true;                                            
                    else
                    {
                        radioButton30.Checked = true;
                        textBox5.Text = myReader.GetString(6);
                    }
                    comboBox3.SelectedValue = myReader.GetInt32(7);
                    comboBox4.SelectedValue = myReader.GetInt32(8);
                    if(myReader.GetString(9).Equals("Tala Raza"))
                        radioButton9.Checked = true;
                    else if(myReader.GetString(9).Equals("Entresaca"))
                        radioButton10.Checked = true;
                    else
                        radioButton8.Checked = true;
                    if(myReader.GetString(10).Equals("Farmi Primero 9000"))
                        radioButton16.Checked = true;
                    else
                        radioButton15.Checked = true;
                    textBox7.Text = myReader.GetInt32(11).ToString();
                    textBox8.Text = myReader.GetInt32(12).ToString();
                    textBox9.Text = myReader.GetInt32(13).ToString();
                    textBox10.Text = myReader.GetDouble(14).ToString();
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

        public void getOrden2(int semana, string orden, string adf, TreeView tree)
        {
            string query = "SELECT * FROM Transferencia WHERE Semana = " + semana + " AND Orden = " + orden + " AND ADF = '" + adf + "' AND Dia = '" + tree.SelectedNode.Parent.Text + "' AND Recorrido = " + tree.SelectedNode.Text;
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
                    if (myReader.GetInt32(5) == 0)
                        radioButton24.Checked = true;
                    else
                        radioButton19.Checked = true;
                    textBox37.Text = myReader.GetInt32(5).ToString();
                    if (myReader.GetString(6).Equals("Teca"))
                        radioButton17.Checked = true;
                    else if (myReader.GetString(6).Equals("Melina"))
                        radioButton18.Checked = true;
                    else
                    {
                        radioButton28.Checked = true;
                        textBox36.Text = myReader.GetString(6);
                    }
                    transp = myReader.GetInt32(8);
                    if (myReader.GetString(10).Equals("Camion 600"))
                        radioButton26.Checked = true;
                    else if (myReader.GetString(10).Equals("Camion Doble Torque"))
                        radioButton25.Checked = true;
                    else
                        radioButton27.Checked = true;
                    textBox41.Text = myReader.GetInt32(11).ToString();
                    textBox40.Text = myReader.GetInt32(12).ToString();
                    textBox39.Text = myReader.GetInt32(13).ToString();
                    textBox38.Text = myReader.GetDouble(14).ToString();
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
                comboBox11.SelectedValue = transp;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string extraido = "";
            if (radioButton10.Checked)
                extraido = "extraidoEntresaca";
            else if (radioButton9.Checked)
                extraido = "extraidoTalaRaza";
            else
                extraido = "extraidoRecuperacion";
            if (transferenciaExiste(semanaFinal, OT, "ADF013", treeView1)) 
            {
                modificarExtraccion(codigo,extraido, textBox10.Text, "-");
                modificarTransferencia(semanaFinal, OT, "ADF013", treeView1);
                agregarLog("Usuario ha modificado el formato ADF013 en la OT #" + OT, userName);
                modificarExtraccion(codigo, extraido, textBox10.Text, "+");
            }
            else
            {
                agregarTransferencia2(semanaFinal, OT, "ADF013", treeView1);
                modificarExtraccion(codigo, extraido, textBox10.Text, "+");
                agregarLog("Usuario ha modificado el formato ADF013 en la OT #" + OT, userName);
            }
            MessageBox.Show("Recorrido # " + treeView1.SelectedNode.Text + " del dia " + treeView1.SelectedNode.Parent.Text + " registrado.");
        }

        public void modificarExtraccion(int lote, string tipo, string volumen, string simbolo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Lotes SET " + tipo + " = (" + tipo + " " + simbolo + " @volumen ) WHERE Codigo = " + lote);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {

                try
                {
                    cmd.Parameters.Add("@voluem", OleDbType.VarChar).Value = volumen;
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

        public void agregarTransferencia(int semana, string orden, string adf, TreeView tree)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Transferencia (Semana,Orden,Dia,Recorrido,FSC,Especie,Maquina,Conductor,motivoRaleo,Trailer,Diametro,Largo,Cantidad,Volumen,ADF) VALUES (@Semana,@Orden,@Dia,@Recorrido,@FSC,@Especie,@Maquina,@Conductor,@motivoRaleo,@Trailer,@Diametro,@Largo,@Cantidad,@Volumen,@ADF)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {

                try
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    cmd.Parameters.Add("@Dia", OleDbType.VarChar).Value = tree.SelectedNode.Parent.Text;
                    cmd.Parameters.Add("@Recorrido", OleDbType.VarChar).Value = tree.SelectedNode.Text;
                    if (textBox37.Text.Equals(""))
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                    else
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox37.Text;
                    if (radioButton18.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Melina";
                    else if (radioButton17.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Teca";
                    else
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = textBox36.Text;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Conductor", OleDbType.VarChar).Value = comboBox11.SelectedValue;
                    cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "N/A";
                    if (radioButton26.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Camion 600";
                    else if (radioButton25.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Camion Doble Torque";
                    else
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Tractomula";
                    cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox41.Text;
                    cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = textBox40.Text;
                    cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox39.Text;
                    cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = textBox38.Text;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
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

        public void agregarTransferencia2(int semana, string orden, string adf, TreeView tree)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("INSERT INTO Transferencia (Semana,Orden,Dia,Recorrido,FSC,Especie,Maquina,Conductor,motivoRaleo,Trailer,Diametro,Largo,Cantidad,Volumen,ADF) VALUES (@Semana,@Orden,@Dia,@Recorrido,@FSC,@Especie,@Maquina,@Conductor,@motivoRaleo,@Trailer,@Diametro,@Largo,@Cantidad,@Volumen,@ADF)");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {

                try
                {
                    cmd.Parameters.Add("@Semana", OleDbType.VarChar).Value = semana;
                    cmd.Parameters.Add("@Orden", OleDbType.VarChar).Value = orden;
                    cmd.Parameters.Add("@Dia", OleDbType.VarChar).Value = tree.SelectedNode.Parent.Text;
                    cmd.Parameters.Add("@Recorrido", OleDbType.VarChar).Value = tree.SelectedNode.Text;
                    if (textBox6.Text.Equals(""))
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                    else
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                    if (radioButton12.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Melina";
                    else if (radioButton13.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Teca";
                    else
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = textBox5.Text;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                    cmd.Parameters.Add("@Conductor", OleDbType.VarChar).Value = comboBox4.SelectedValue;
                    if (radioButton10.Checked)
                        cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "Entresaca";
                    else if (radioButton9.Checked)
                        cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "Tala Raza";
                    else
                        cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "Recuperacion de Material";
                    if (radioButton16.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Farmi Primero 9000";
                    else if (radioButton15.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Pfanzelt 15100";
                    cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox7.Text;
                    cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = textBox8.Text;
                    cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox9.Text;
                    cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = textBox10.Text;
                    cmd.Parameters.Add("@ADF", OleDbType.VarChar).Value = adf;
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

        public void modificarTransferencia(int semana, string orden,string adf, TreeView tree)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Transferencia SET FSC=@FSC,Especie=@Especie,Maquina=@Maquina,Conductor=@Conductor,motivoRaleo=@motivoRaleo,Trailer=@Trailer,Diametro=@Diametro,Largo=@Largo,Cantidad=@Cantidad,Volumen=@Volumen WHERE Semana = " + semana + " AND Dia = '" + tree.SelectedNode.Parent.Text + "' AND Recorrido = " + tree.SelectedNode.Text + " AND Orden = " + orden + " AND ADF = '" + adf + "'");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {

                try
                {
                    if (textBox6.Text.Equals(""))
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                    else
                    cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox6.Text;
                    if(radioButton12.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Melina";
                    else if (radioButton13.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Teca";
                    else
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = textBox5.Text;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = comboBox3.SelectedValue;
                    cmd.Parameters.Add("@Conductor", OleDbType.VarChar).Value = comboBox4.SelectedValue;
                    if (radioButton10.Checked)
                        cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "Entresaca";
                    else if (radioButton9.Checked)
                        cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "Tala Raza";
                    else
                        cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "Recuperacion de Material";
                    if (radioButton16.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Farmi Primero 9000";
                    else if (radioButton15.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Pfanzelt 15100";
                    cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox7.Text;
                    cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = textBox8.Text;
                    cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox9.Text;
                    cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = textBox10.Text;
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

        public void modificarTransferencia2(int semana, string orden, string adf, TreeView tree)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand("UPDATE Transferencia SET FSC=@FSC,Especie=@Especie,Maquina=@Maquina,Conductor=@Conductor,motivoRaleo=@motivoRaleo,Trailer=@Trailer,Diametro=@Diametro,Largo=@Largo,Cantidad=@Cantidad,Volumen=@Volumen WHERE Semana = " + semana + " AND Dia = '" + tree.SelectedNode.Parent.Text + "' AND Recorrido = " + tree.SelectedNode.Text + " AND Orden = " + orden + " AND ADF = '" + adf + "'");
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {

                try
                {
                    if (textBox37.Text.Equals(""))
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = 0;
                    else
                        cmd.Parameters.Add("@FSC", OleDbType.VarChar).Value = textBox37.Text;
                    if (radioButton18.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Melina";
                    else if (radioButton17.Checked)
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = "Teca";
                    else
                        cmd.Parameters.Add("@Especie", OleDbType.VarChar).Value = textBox36.Text;
                    cmd.Parameters.Add("@Maquina", OleDbType.VarChar).Value = 0;
                    cmd.Parameters.Add("@Conductor", OleDbType.VarChar).Value = comboBox11.SelectedValue;
                    cmd.Parameters.Add("@motivoRaleo", OleDbType.VarChar).Value = "N/A";
                    if (radioButton26.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Camion 600";
                    else if (radioButton25.Checked)
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Camion Doble Torque";
                    else
                        cmd.Parameters.Add("@Trailer", OleDbType.VarChar).Value = "Tractomula";
                    cmd.Parameters.Add("@Diametro", OleDbType.VarChar).Value = textBox41.Text;
                    cmd.Parameters.Add("@Largo", OleDbType.VarChar).Value = textBox40.Text;
                    cmd.Parameters.Add("@Cantidad", OleDbType.VarChar).Value = textBox39.Text;
                    cmd.Parameters.Add("@Volumen", OleDbType.VarChar).Value = textBox38.Text;
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

        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView2.SelectedNode.Parent != null)
            {
                label75.Text = "Dia: " + treeView2.SelectedNode.Parent.Text + " Recorrido #: " + treeView2.SelectedNode.Text;
                textBox41.Text = "";
                textBox40.Text = "";
                textBox39.Text = "";
                textBox38.Text = "";
                getOrden2(semanaFinal, OT, "ADF015", treeView2);
            }
        }

        private void dataGridView13_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView13.Columns.Count > 5)
            {
                dataGridView13.Rows[dataGridView13.Rows.Count - 1].Cells[0].Value = dataGridView13.Rows.Count;
                for (int i = 0; i < 7; i++)
                {
                    dataGridView13.Rows[dataGridView13.Rows.Count - 1].Cells[i+5].Value = 0;      
                }                
            }
        }

        private void dataGridView13_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView13.Columns.Count > 10)
            {
                for (int i = 0; i < dataGridView13.Rows.Count; i++)
                {
                    if (dataGridView13.Rows[i].Cells[2].FormattedValue != null)
                    {
                        if (!dataGridView13.Rows[i].Cells[2].FormattedValue.ToString().Equals(""))
                        {
                            string cedula = getCedula(Int32.Parse(dataGridView13.Rows[i].Cells[2].Value.ToString()));
                            dataGridView13.Rows[i].Cells[3].Value = cedula;
                        }                        
                    }
                }
            }
            ContadorSemanal(dataGridView13, 5, 11);
        }

        private void dataGridView13_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView13);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (semillaExiste(semanaFinal, OT) == false)
            {
                crearSemilla(semanaFinal, OT,dataGridView13);                
                agregarADF(semanaFinal, OT);
                agregarLog("Usuario ha modificado el formato ADF010 en la OT #" + OT, userName);
            }
            else
            {
                eliminarADF(semanaFinal, OT);
                modificarSemilla(semanaFinal, OT, dataGridView13);
                agregarADF(semanaFinal, OT);
                agregarLog("Usuario ha modificado el formato ADF009 en la OT #" + OT, userName);
            }
            MessageBox.Show("Control de Insumos Recoleccion de Semilla registrado.");
        }

        private void dataGridView13_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (comboBox5.SelectedItem != null)
            {
                if (mantExiste(semanaFinal, OT, "ADF016") == false)
                {
                    crearADF(semanaFinal, OT, "ADF016", dataGridView15, Maquinarias[0], comboBox5, dateTimePicker1, "Motor");
                    crearADF(semanaFinal, OT, "ADF016", dataGridView16, Maquinarias[0], comboBox5, dateTimePicker1, "Hidraulico");
                    crearADF(semanaFinal, OT, "ADF016", dataGridView17, Maquinarias[0], comboBox5, dateTimePicker1, "Diferenciales");
                    crearADF(semanaFinal, OT, "ADF016", dataGridView18, Maquinarias[0], comboBox5, dateTimePicker1, "Adicional");
                    cargarADFMant(semanaFinal, OT, "ADF016", dataGridView15, "Motor");
                    cargarADFMant(semanaFinal, OT, "ADF016", dataGridView16, "Hidraulico");
                    cargarADFMant(semanaFinal, OT, "ADF016", dataGridView17, "Diferenciales");
                    cargarADFMantDetalle(semanaFinal, OT, "ADF016", dataGridView18, comboBox5, dateTimePicker1);
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "+");
                    //eliminarADFMant(semanaFinal, OT);
                    agregarLog("Usuario ha modificado el formato ADF016 en la OT #" + OT, userName);
                }
                else
                {
                    //agregarADFMant(semanaFinal, OT);
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "-");
                    modificarADF(semanaFinal, OT, "ADF016", dataGridView15, Maquinarias[0], comboBox5, dateTimePicker1, "Motor");
                    modificarADF(semanaFinal, OT, "ADF016", dataGridView16, Maquinarias[0], comboBox5, dateTimePicker1, "Hidraulico");
                    modificarADF(semanaFinal, OT, "ADF016", dataGridView17, Maquinarias[0], comboBox5, dateTimePicker1, "Diferenciales");
                    modificarADF(semanaFinal, OT, "ADF016", dataGridView18, Maquinarias[0], comboBox5, dateTimePicker1, "Adicional");
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "+");
                    //eliminarADFMant(semanaFinal, OT);
                    agregarLog("Usuario ha modificado el formato ADF016 en la OT #" + OT, userName);
                }
                modificarHorometro(Maquinarias[0], Int32.Parse(textBox11.Text));
                MessageBox.Show("Mantenimiento TigerCat registrado.");
            }
            else
            {
                MessageBox.Show("Favor seleccionar quien realizo el mantenimiento.");
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if(comboBox7.SelectedValue != null)
            {
                if (mantExiste(semanaFinal, OT, "ADF018") == false)
                {
                    crearADF(semanaFinal, OT, "ADF018", dataGridView26, Maquinarias[0], comboBox7, dateTimePicker3, "Motor");
                    crearADF(semanaFinal, OT, "ADF018", dataGridView25, Maquinarias[0], comboBox7, dateTimePicker3, "Hidraulico");
                    crearADF(semanaFinal, OT, "ADF018", dataGridView24, Maquinarias[0], comboBox7, dateTimePicker3, "Diferenciales");
                    crearADF(semanaFinal, OT, "ADF018", dataGridView23, Maquinarias[0], comboBox7, dateTimePicker3, "Adicional");
                    cargarADFMant(semanaFinal, OT, "ADF018", dataGridView26, "Motor");
                    cargarADFMant(semanaFinal, OT, "ADF018", dataGridView25, "Hidraulico");
                    cargarADFMant(semanaFinal, OT, "ADF018", dataGridView24, "Diferenciales");
                    cargarADFMantDetalle(semanaFinal, OT, "ADF018", dataGridView23, comboBox7, dateTimePicker3);
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "+");
                    //eliminarADFMant(semanaFinal, OT);
                    agregarLog("Usuario ha modificado el formato ADF018 en la OT #" + OT, userName);
                }
                else
                {
                    //agregarADFMant(semanaFinal, OT);
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "-");
                    modificarADF(semanaFinal, OT, "ADF018", dataGridView26, Maquinarias[0], comboBox7, dateTimePicker3, "Motor");
                    modificarADF(semanaFinal, OT, "ADF018", dataGridView25, Maquinarias[0], comboBox7, dateTimePicker3, "Hidraulico");
                    modificarADF(semanaFinal, OT, "ADF018", dataGridView24, Maquinarias[0], comboBox7, dateTimePicker3, "Diferenciales");
                    modificarADF(semanaFinal, OT, "ADF018", dataGridView23, Maquinarias[0], comboBox7, dateTimePicker3, "Adicional");
                    controlCosto(Int32.Parse(OT), getMantCosto(semanaFinal, Int32.Parse(OT)), "+");
                    //eliminarADFMant(semanaFinal, OT);
                    agregarLog("Usuario ha modificado el formato ADF018 en la OT #" + OT, userName);
                }
                modificarHorometro(Maquinarias[0], Int32.Parse(textBox21.Text));
                MessageBox.Show("Mantenimiento General registrado.");
            }
            else
            {
                MessageBox.Show("Favor seleccionar quien realizo el mantenimiento.");
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (ADFExiste(semanaFinal, OT, "ADF011") == false)
            {
                crearADF3(semanaFinal, OT, "ADF011", dataGridView14);
                //agregarADF(semanaFinal, OT, "ADF011");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF011"), "-");            
                agregarLog("Usuario ha modificado el formato ADF011 en la OT #" + OT, userName);
            }
            else
            {
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF011"), "-");            
                //agregarADF(semanaFinal, OT, "ADF011");
                modificarADF011(semanaFinal, OT, "ADF011", dataGridView14);
                //eliminarADF(semanaFinal, OT, "ADF011");
                controlCosto(Int32.Parse(OT), getControlCosto(semanaFinal, Int32.Parse(OT), "ADF011"), "+");            
                agregarLog("Usuario ha modificado el formato ADF011 en la OT #" + OT, userName);
            }
            MessageBox.Show("Control de Actividades Semanales registrada.");
        }

        private void dataGridView12_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
        }

        private void dataGridView14_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView14.Columns.Count > 5)
            {
                dataGridView14.Rows[dataGridView14.Rows.Count - 1].Cells[0].Value = dataGridView14.Rows.Count;
            }
        }

        private void dataGridView14_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView14.Columns.Count > 10)
            {
                if (termino == 1)
                {
                    for (int i = 0; i < dataGridView14.Rows.Count; i++)
                    {
                        if (dataGridView14.Rows[i].Cells[3].FormattedValue != null && !dataGridView14.Rows[i].Cells[3].FormattedValue.Equals(""))
                        {
                            formato(dataGridView14, getUnidad((int)dataGridView14.Rows[i].Cells[3].Value), i);
                        }
                    }
                    Contador(dataGridView14, 5, 17);
                }
            }
        }

        private void dataGridView14_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dirtyCell(dataGridView14);
        }

        public int getControlCosto(int semana, int orden, string adf) {

            string query = "SELECT  SUM((c.Lunes+c.Martes+c.Miercoles+c.Jueves+c.Viernes+c.Sabado)*i.Costo_Unitario) FROM (Control AS c INNER JOIN historicoOrdenes AS h ON c.Orden = h.ID) INNER JOIN Insumos AS i ON c.Modelo = i.ID WHERE h.ID = " + orden + " AND c.Semana = " + semana + " AND c.ADF = '" + adf +"'";
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            int costo = 0;
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    if(!myReader.GetValue(0).ToString().Equals(""))
                        costo = Int32.Parse(myReader.GetValue(0).ToString());
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return costo;
        }

        public int getMantCosto(int semana, int orden)
        {
            string query = "SELECT  SUM((c.Cantidad)*i.Costo_Unitario) FROM ( formatoMantenimiento AS c INNER JOIN historicoOrdenes AS h ON c.Orden = h.ID) INNER JOIN Insumos AS i ON c.Insumo = i.ID WHERE h.ID = " + orden + " AND c.Semana = " + semana;
            //Ejecutar el query y llenar el GridView.
            conn.ConnectionString = connectionString;
            OleDbCommand cmd = new OleDbCommand(query, conn);
            cmd.Connection = conn;
            conn.Open();
            int costo = 0;
            OleDbDataReader myReader = cmd.ExecuteReader();
            try
            {
                if (myReader.Read())
                {
                    if (!myReader.GetValue(0).ToString().Equals(""))
                        costo = Int32.Parse(myReader.GetValue(0).ToString());
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return costo;
        }

        public void controlCosto(int orden, int costo, string simbolo)
        {
            conn.ConnectionString = connectionString;
            OleDbCommand cmd;
            cmd = new OleDbCommand("UPDATE historicoOrdenes SET costoFinal = costoFinal " + simbolo + " @costo WHERE ID = " + orden);
            cmd.Connection = conn;
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@costo", OleDbType.VarChar).Value = costo;
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmOrdenFormatos newFrm = new frmOrdenFormatos(OT, semanaActual-1,tipousuario);
            this.Hide();
            newFrm.ShowDialog();
            this.Close();
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            Variables.cargar(comboBox8, "SELECT * FROM Insumos WHERE Clase = '" + comboBox12.Text + "'", "Modelo");
            comboBox10.Items.Clear();
            Variables.cargar2(comboBox10, "SELECT Marca FROM Insumos WHERE Clase = '" + comboBox12.Text + "' GROUP BY Marca", "Marca");
        }

        private void comboBox10_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (comboBox12.Text.Equals(""))
                Variables.cargar(comboBox8, "SELECT * FROM Insumos WHERE Marca = '" + comboBox10.Text + "'", "Modelo");
            else
                Variables.cargar(comboBox8, "SELECT * FROM Insumos WHERE Marca = '" + comboBox10.Text + "' AND Clase = '" + comboBox12.Text + "'", "Modelo");
            //Variables.cargar(comboBox8, "SELECT * FROM Insumos WHERE Marca = '" + comboBox10.Text + "'", "Modelo");
            //Variables.cargar2(comboBox12, "SELECT Clase FROM Insumos WHERE Marca = '" + comboBox10.Text + "' GROUP BY Clase", "Clase");
        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            if (comboBox8.SelectedItem != null)
            {
                dataGridView14.Rows.Add();
                dataGridView14.Rows[dataGridView14.Rows.Count - 1].Cells[3].Value = comboBox8.SelectedValue;
            }
        }

        public double getVolumen(double diametro, double largo, int cantidad, string especie, string trailer, string raleo, string adf)
        {
            double volumen = 0;
            string query = "SELECT Volumen FROM volumenCalculado WHERE Diametro = " + diametro.ToString().Replace(",", ".") + " AND Largo = " + largo.ToString().Replace(",", ".") +" AND Cantidad = " + cantidad + " AND Especie = '" + especie + "' AND Trailer = '" + trailer + "' and Raleo = '" + raleo + "'";
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
                    volumen = double.Parse(myReader.GetValue(0).ToString());
                }
                else
                {
                    volumen = 0;
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // always call Close when done reading.
                conn.Close();
            }
            return volumen;
        }

        private void textBox38_Enter(object sender, EventArgs e)
        {
            if (!textBox41.Text.Equals("") && !textBox40.Text.Equals("") && !textBox39.Text.Equals(""))
            {
                string raleo = "", especie = "", trailer = "";
                raleo = "N/A";
                especie = getEspecie(radioButton18, radioButton17, radioButton28, textBox36);
                if (radioButton26.Checked)
                    trailer = "Camion 600";
                else if (radioButton27.Checked)
                    trailer = "Camion Doble Torque";
                else
                    trailer = "Tractomula";
                textBox38.Text = getVolumen(double.Parse(textBox41.Text), double.Parse(textBox40.Text), Int32.Parse(textBox39.Text), especie, trailer, raleo, "ADF012").ToString();
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (transferenciaExiste(semanaFinal, OT, "ADF015", treeView2))
            {
                modificarTransferencia2(semanaFinal, OT, "ADF015", treeView2);
                agregarLog("Usuario ha modificado el formato ADF015 en la OT #" + OT, userName);

            }
            else
            {
                agregarTransferencia(semanaFinal, OT, "ADF015", treeView2);
                agregarLog("Usuario ha modificado el formato ADF015 en la OT #" + OT, userName);

            }
            MessageBox.Show("Recorrido # " + treeView2.SelectedNode.Text + " del dia " + treeView2.SelectedNode.Parent.Text + " registrado.");
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            if (!textBox7.Text.Equals("") && !textBox8.Text.Equals("") && !textBox9.Text.Equals(""))
            {
                string raleo = "", especie = "", trailer = "";
                raleo = getRaleo(radioButton10, radioButton9, radioButton8);
                especie = getEspecie(radioButton12, radioButton11, radioButton30, textBox5);
                if (radioButton16.Checked)
                    trailer = "Farmi Primero 9000";
                else
                    trailer = "Pfanzelt 15100";
                textBox10.Text = getVolumen(double.Parse(textBox7.Text), double.Parse(textBox8.Text), Int32.Parse(textBox9.Text), especie, trailer, raleo, "ADF012").ToString();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmOrdenFormatos newFrm = new frmOrdenFormatos(OT, semanaActual + 1,tipousuario);
            this.Hide();
            newFrm.ShowDialog();
            this.Close();
        }

    }
}
