namespace Aplicativo
{
    partial class frmEmpleados
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCargos = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.btnPredial = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.textTerminado = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.dateTimePicker3 = new System.Windows.Forms.DateTimePicker();
            this.label15 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label14 = new System.Windows.Forms.Label();
            this.Todos = new System.Windows.Forms.RadioButton();
            this.btnBuscar = new System.Windows.Forms.Button();
            this.Contratista = new System.Windows.Forms.RadioButton();
            this.fijoTH = new System.Windows.Forms.RadioButton();
            this.fijoFHCTH = new System.Windows.Forms.RadioButton();
            this.label13 = new System.Windows.Forms.Label();
            this.txtMaquinaria = new System.Windows.Forms.ComboBox();
            this.btnReiniciar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnModificar = new System.Windows.Forms.Button();
            this.btnAgregar = new System.Windows.Forms.Button();
            this.txtBotas = new System.Windows.Forms.ComboBox();
            this.txtPantalon = new System.Windows.Forms.ComboBox();
            this.txtCamisa = new System.Windows.Forms.ComboBox();
            this.txtCargo = new System.Windows.Forms.ComboBox();
            this.txtDepartamento = new System.Windows.Forms.ComboBox();
            this.txtCelular = new System.Windows.Forms.TextBox();
            this.txtDir = new System.Windows.Forms.TextBox();
            this.txtSangre = new System.Windows.Forms.ComboBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.txtCedula = new System.Windows.Forms.TextBox();
            this.txtApellido = new System.Windows.Forms.TextBox();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCargos
            // 
            this.btnCargos.Location = new System.Drawing.Point(1006, 203);
            this.btnCargos.Name = "btnCargos";
            this.btnCargos.Size = new System.Drawing.Size(122, 22);
            this.btnCargos.TabIndex = 0;
            this.btnCargos.Text = "Cargos Laborales";
            this.btnCargos.UseVisualStyleBackColor = true;
            this.btnCargos.Click += new System.EventHandler(this.btnCargos_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 30);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(1492, 344);
            this.dataGridView1.TabIndex = 21;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.btnCerrar);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.label17);
            this.groupBox1.Controls.Add(this.textTerminado);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.dateTimePicker3);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.dateTimePicker2);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.Todos);
            this.groupBox1.Controls.Add(this.btnBuscar);
            this.groupBox1.Controls.Add(this.Contratista);
            this.groupBox1.Controls.Add(this.fijoTH);
            this.groupBox1.Controls.Add(this.fijoFHCTH);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.txtMaquinaria);
            this.groupBox1.Controls.Add(this.btnReiniciar);
            this.groupBox1.Controls.Add(this.btnEliminar);
            this.groupBox1.Controls.Add(this.btnModificar);
            this.groupBox1.Controls.Add(this.btnAgregar);
            this.groupBox1.Controls.Add(this.txtBotas);
            this.groupBox1.Controls.Add(this.txtPantalon);
            this.groupBox1.Controls.Add(this.btnCargos);
            this.groupBox1.Controls.Add(this.txtCamisa);
            this.groupBox1.Controls.Add(this.txtCargo);
            this.groupBox1.Controls.Add(this.txtDepartamento);
            this.groupBox1.Controls.Add(this.txtCelular);
            this.groupBox1.Controls.Add(this.txtDir);
            this.groupBox1.Controls.Add(this.txtSangre);
            this.groupBox1.Controls.Add(this.dateTimePicker1);
            this.groupBox1.Controls.Add(this.txtCedula);
            this.groupBox1.Controls.Add(this.txtApellido);
            this.groupBox1.Controls.Add(this.txtNombre);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 380);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1492, 279);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Agregar Empleado";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBox12);
            this.groupBox2.Controls.Add(this.btnPredial);
            this.groupBox2.Controls.Add(this.pictureBox1);
            this.groupBox2.Location = new System.Drawing.Point(32, 18);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(161, 224);
            this.groupBox2.TabIndex = 45;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Imagen Empelado";
            // 
            // textBox12
            // 
            this.textBox12.Location = new System.Drawing.Point(5, 170);
            this.textBox12.Margin = new System.Windows.Forms.Padding(2);
            this.textBox12.Name = "textBox12";
            this.textBox12.ReadOnly = true;
            this.textBox12.Size = new System.Drawing.Size(150, 20);
            this.textBox12.TabIndex = 62;
            // 
            // btnPredial
            // 
            this.btnPredial.Location = new System.Drawing.Point(5, 195);
            this.btnPredial.Name = "btnPredial";
            this.btnPredial.Size = new System.Drawing.Size(150, 22);
            this.btnPredial.TabIndex = 63;
            this.btnPredial.Text = "Examinar";
            this.btnPredial.UseVisualStyleBackColor = true;
            this.btnPredial.Click += new System.EventHandler(this.btnPredial_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(5, 16);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(150, 150);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(6, 247);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(1480, 24);
            this.btnCerrar.TabIndex = 34;
            this.btnCerrar.Text = "Cerrar Ventana";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(1203, 157);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(57, 20);
            this.textBox2.TabIndex = 44;
            this.textBox2.Text = "0";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(1116, 159);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(81, 13);
            this.label18.TabIndex = 43;
            this.label18.Text = "Dias Laborados";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(1203, 120);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(57, 20);
            this.textBox1.TabIndex = 42;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(1114, 123);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(85, 13);
            this.label17.TabIndex = 41;
            this.label17.Text = "Ubicación Fisica";
            // 
            // textTerminado
            // 
            this.textTerminado.FormattingEnabled = true;
            this.textTerminado.Items.AddRange(new object[] {
            "Si",
            "No"});
            this.textTerminado.Location = new System.Drawing.Point(695, 162);
            this.textTerminado.Name = "textTerminado";
            this.textTerminado.Size = new System.Drawing.Size(68, 21);
            this.textTerminado.TabIndex = 40;
            this.textTerminado.TextChanged += new System.EventHandler(this.textTerminado_TextChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(586, 166);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(103, 13);
            this.label16.TabIndex = 39;
            this.label16.Text = "Empleado terminado";
            // 
            // dateTimePicker3
            // 
            this.dateTimePicker3.CustomFormat = "";
            this.dateTimePicker3.Location = new System.Drawing.Point(885, 163);
            this.dateTimePicker3.Name = "dateTimePicker3";
            this.dateTimePicker3.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker3.TabIndex = 38;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(766, 166);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(113, 13);
            this.label15.TabIndex = 37;
            this.label15.Text = "Fecha de Terminación";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.CustomFormat = "";
            this.dateTimePicker2.Location = new System.Drawing.Point(377, 162);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker2.TabIndex = 36;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(281, 166);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(90, 13);
            this.label14.TabIndex = 35;
            this.label14.Text = "Fecha de Ingreso";
            // 
            // Todos
            // 
            this.Todos.AutoSize = true;
            this.Todos.Location = new System.Drawing.Point(487, 121);
            this.Todos.Name = "Todos";
            this.Todos.Size = new System.Drawing.Size(55, 17);
            this.Todos.TabIndex = 34;
            this.Todos.TabStop = true;
            this.Todos.Text = "Todos";
            this.Todos.UseVisualStyleBackColor = true;
            // 
            // btnBuscar
            // 
            this.btnBuscar.Location = new System.Drawing.Point(620, 203);
            this.btnBuscar.Name = "btnBuscar";
            this.btnBuscar.Size = new System.Drawing.Size(122, 22);
            this.btnBuscar.TabIndex = 33;
            this.btnBuscar.Text = "Buscar";
            this.btnBuscar.UseVisualStyleBackColor = true;
            this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
            // 
            // Contratista
            // 
            this.Contratista.AutoSize = true;
            this.Contratista.Location = new System.Drawing.Point(688, 121);
            this.Contratista.Name = "Contratista";
            this.Contratista.Size = new System.Drawing.Size(75, 17);
            this.Contratista.TabIndex = 32;
            this.Contratista.TabStop = true;
            this.Contratista.Text = "Contratista";
            this.Contratista.UseVisualStyleBackColor = true;
            // 
            // fijoTH
            // 
            this.fijoTH.AutoSize = true;
            this.fijoTH.Location = new System.Drawing.Point(626, 121);
            this.fijoTH.Name = "fijoTH";
            this.fijoTH.Size = new System.Drawing.Size(57, 17);
            this.fijoTH.TabIndex = 31;
            this.fijoTH.TabStop = true;
            this.fijoTH.Text = "Fijo Th";
            this.fijoTH.UseVisualStyleBackColor = true;
            // 
            // fijoFHCTH
            // 
            this.fijoFHCTH.AutoSize = true;
            this.fijoFHCTH.Location = new System.Drawing.Point(548, 121);
            this.fijoFHCTH.Name = "fijoFHCTH";
            this.fijoFHCTH.Size = new System.Drawing.Size(72, 17);
            this.fijoFHCTH.TabIndex = 30;
            this.fijoFHCTH.TabStop = true;
            this.fijoFHCTH.Text = "Fijo FCTH";
            this.fijoFHCTH.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(782, 125);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(59, 13);
            this.label13.TabIndex = 29;
            this.label13.Text = "Maquinaria";
            // 
            // txtMaquinaria
            // 
            this.txtMaquinaria.FormattingEnabled = true;
            this.txtMaquinaria.Location = new System.Drawing.Point(847, 120);
            this.txtMaquinaria.Name = "txtMaquinaria";
            this.txtMaquinaria.Size = new System.Drawing.Size(251, 21);
            this.txtMaquinaria.TabIndex = 28;
            // 
            // btnReiniciar
            // 
            this.btnReiniciar.Location = new System.Drawing.Point(878, 203);
            this.btnReiniciar.Name = "btnReiniciar";
            this.btnReiniciar.Size = new System.Drawing.Size(122, 22);
            this.btnReiniciar.TabIndex = 27;
            this.btnReiniciar.Text = "Reiniciar Tablero";
            this.btnReiniciar.UseVisualStyleBackColor = true;
            this.btnReiniciar.Click += new System.EventHandler(this.btnReiniciar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Location = new System.Drawing.Point(749, 203);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(122, 22);
            this.btnEliminar.TabIndex = 26;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnModificar
            // 
            this.btnModificar.Location = new System.Drawing.Point(494, 203);
            this.btnModificar.Name = "btnModificar";
            this.btnModificar.Size = new System.Drawing.Size(122, 22);
            this.btnModificar.TabIndex = 25;
            this.btnModificar.Text = "Modificar";
            this.btnModificar.UseVisualStyleBackColor = true;
            this.btnModificar.Click += new System.EventHandler(this.btnModificar_Click);
            // 
            // btnAgregar
            // 
            this.btnAgregar.Location = new System.Drawing.Point(365, 203);
            this.btnAgregar.Name = "btnAgregar";
            this.btnAgregar.Size = new System.Drawing.Size(122, 22);
            this.btnAgregar.TabIndex = 24;
            this.btnAgregar.Text = "Agregar";
            this.btnAgregar.UseVisualStyleBackColor = true;
            this.btnAgregar.Click += new System.EventHandler(this.btnAgregar_Click);
            // 
            // txtBotas
            // 
            this.txtBotas.FormattingEnabled = true;
            this.txtBotas.Items.AddRange(new object[] {
            "36",
            "37",
            "38",
            "39",
            "40",
            "41",
            "42",
            "43",
            "44",
            "45",
            "46",
            "47",
            "48"});
            this.txtBotas.Location = new System.Drawing.Point(1203, 94);
            this.txtBotas.Name = "txtBotas";
            this.txtBotas.Size = new System.Drawing.Size(57, 21);
            this.txtBotas.TabIndex = 23;
            // 
            // txtPantalon
            // 
            this.txtPantalon.FormattingEnabled = true;
            this.txtPantalon.Items.AddRange(new object[] {
            "26",
            "28",
            "30",
            "32",
            "34",
            "36",
            "38",
            "40"});
            this.txtPantalon.Location = new System.Drawing.Point(1203, 67);
            this.txtPantalon.Name = "txtPantalon";
            this.txtPantalon.Size = new System.Drawing.Size(57, 21);
            this.txtPantalon.TabIndex = 22;
            // 
            // txtCamisa
            // 
            this.txtCamisa.FormattingEnabled = true;
            this.txtCamisa.Items.AddRange(new object[] {
            "S",
            "M",
            "L",
            "XL"});
            this.txtCamisa.Location = new System.Drawing.Point(1203, 40);
            this.txtCamisa.Name = "txtCamisa";
            this.txtCamisa.Size = new System.Drawing.Size(57, 21);
            this.txtCamisa.TabIndex = 21;
            // 
            // txtCargo
            // 
            this.txtCargo.FormattingEnabled = true;
            this.txtCargo.Location = new System.Drawing.Point(884, 93);
            this.txtCargo.Name = "txtCargo";
            this.txtCargo.Size = new System.Drawing.Size(214, 21);
            this.txtCargo.TabIndex = 20;
            // 
            // txtDepartamento
            // 
            this.txtDepartamento.FormattingEnabled = true;
            this.txtDepartamento.Location = new System.Drawing.Point(884, 66);
            this.txtDepartamento.Name = "txtDepartamento";
            this.txtDepartamento.Size = new System.Drawing.Size(214, 21);
            this.txtDepartamento.TabIndex = 19;
            // 
            // txtCelular
            // 
            this.txtCelular.Location = new System.Drawing.Point(563, 93);
            this.txtCelular.Name = "txtCelular";
            this.txtCelular.Size = new System.Drawing.Size(200, 20);
            this.txtCelular.TabIndex = 18;
            // 
            // txtDir
            // 
            this.txtDir.Location = new System.Drawing.Point(563, 66);
            this.txtDir.Name = "txtDir";
            this.txtDir.Size = new System.Drawing.Size(200, 20);
            this.txtDir.TabIndex = 17;
            // 
            // txtSangre
            // 
            this.txtSangre.FormattingEnabled = true;
            this.txtSangre.Items.AddRange(new object[] {
            "O-",
            "O+",
            "A-",
            "A+",
            "B-",
            "B+",
            "AB-",
            "AB+"});
            this.txtSangre.Location = new System.Drawing.Point(884, 40);
            this.txtSangre.Name = "txtSangre";
            this.txtSangre.Size = new System.Drawing.Size(162, 21);
            this.txtSangre.TabIndex = 16;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "";
            this.dateTimePicker1.Location = new System.Drawing.Point(563, 40);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 15;
            // 
            // txtCedula
            // 
            this.txtCedula.Location = new System.Drawing.Point(301, 93);
            this.txtCedula.Name = "txtCedula";
            this.txtCedula.Size = new System.Drawing.Size(142, 20);
            this.txtCedula.TabIndex = 14;
            // 
            // txtApellido
            // 
            this.txtApellido.Location = new System.Drawing.Point(301, 66);
            this.txtApellido.Name = "txtApellido";
            this.txtApellido.Size = new System.Drawing.Size(142, 20);
            this.txtApellido.TabIndex = 13;
            // 
            // txtNombre
            // 
            this.txtNombre.Location = new System.Drawing.Point(301, 40);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(142, 20);
            this.txtNombre.TabIndex = 12;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(1137, 97);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(60, 13);
            this.label12.TabIndex = 11;
            this.label12.Text = "Talla Botas";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(1131, 43);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(67, 13);
            this.label11.TabIndex = 10;
            this.label11.Text = "Talla Camisa";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(1122, 69);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(75, 13);
            this.label10.TabIndex = 9;
            this.label10.Text = "Talla Pantalon";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(844, 97);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(35, 13);
            this.label9.TabIndex = 8;
            this.label9.Text = "Cargo";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(805, 69);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(74, 13);
            this.label8.TabIndex = 7;
            this.label8.Text = "Departamento";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(506, 69);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(52, 13);
            this.label7.TabIndex = 6;
            this.label7.Text = "Direccion";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(518, 97);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 13);
            this.label6.TabIndex = 5;
            this.label6.Text = "Celular";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(770, 43);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(109, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Grupo Sanguineo RH";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(449, 43);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(108, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Fecha de Nacimiento";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(250, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Cedula";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(242, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Apellidos";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(242, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Nombres";
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(1200, 14);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(84, 13);
            this.linkLabel1.TabIndex = 72;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Exportar a Excel";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // frmEmpleados
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1516, 671);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "frmEmpleados";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Empleados";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCargos;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtDir;
        private System.Windows.Forms.ComboBox txtSangre;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.TextBox txtCedula;
        private System.Windows.Forms.TextBox txtApellido;
        private System.Windows.Forms.TextBox txtNombre;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCelular;
        private System.Windows.Forms.Button btnReiniciar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnModificar;
        private System.Windows.Forms.Button btnAgregar;
        private System.Windows.Forms.ComboBox txtBotas;
        private System.Windows.Forms.ComboBox txtPantalon;
        private System.Windows.Forms.ComboBox txtCamisa;
        private System.Windows.Forms.ComboBox txtCargo;
        private System.Windows.Forms.ComboBox txtDepartamento;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox txtMaquinaria;
        private System.Windows.Forms.RadioButton Contratista;
        private System.Windows.Forms.RadioButton fijoTH;
        private System.Windows.Forms.RadioButton fijoFHCTH;
        private System.Windows.Forms.Button btnBuscar;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.RadioButton Todos;
        private System.Windows.Forms.DateTimePicker dateTimePicker3;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.ComboBox textTerminado;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textBox12;
        private System.Windows.Forms.Button btnPredial;
    }
}