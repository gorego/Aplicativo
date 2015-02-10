namespace Aplicativo
{
    partial class frmUsuarios
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
            this.gridSupervisor = new System.Windows.Forms.DataGridView();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtEmpleado = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.gridSupervisor)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // gridSupervisor
            // 
            this.gridSupervisor.AllowUserToAddRows = false;
            this.gridSupervisor.AllowUserToDeleteRows = false;
            this.gridSupervisor.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridSupervisor.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.gridSupervisor.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridSupervisor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.gridSupervisor.Location = new System.Drawing.Point(12, 12);
            this.gridSupervisor.Name = "gridSupervisor";
            this.gridSupervisor.ReadOnly = true;
            this.gridSupervisor.RowHeadersVisible = false;
            this.gridSupervisor.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.gridSupervisor.Size = new System.Drawing.Size(745, 292);
            this.gridSupervisor.TabIndex = 2;
            this.gridSupervisor.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridSupervisor_CellClick);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(12, 481);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(745, 30);
            this.btnCerrar.TabIndex = 12;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.btnEliminar);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.txtEmpleado);
            this.groupBox1.Location = new System.Drawing.Point(12, 310);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(745, 165);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Agregar Usuario";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(127, 129);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(485, 30);
            this.button4.TabIndex = 23;
            this.button4.Text = "Dar Permiso Fijo";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(480, 89);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(95, 34);
            this.button3.TabIndex = 22;
            this.button3.Text = "Reiniciar Tablero";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(64, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Contraseña";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(131, 59);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(124, 20);
            this.textBox2.TabIndex = 17;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(82, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Usuario";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(131, 33);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(124, 20);
            this.textBox1.TabIndex = 15;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(379, 89);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(95, 34);
            this.button2.TabIndex = 14;
            this.button2.Text = "Modificar Usuario";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Location = new System.Drawing.Point(278, 89);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(95, 34);
            this.btnEliminar.TabIndex = 10;
            this.btnEliminar.Text = "Eliminar Usuario";
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(317, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Empleado";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(177, 89);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(95, 34);
            this.button1.TabIndex = 12;
            this.button1.Text = "Agregar Usuario";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtEmpleado
            // 
            this.txtEmpleado.FormattingEnabled = true;
            this.txtEmpleado.Location = new System.Drawing.Point(378, 33);
            this.txtEmpleado.Name = "txtEmpleado";
            this.txtEmpleado.Size = new System.Drawing.Size(234, 21);
            this.txtEmpleado.TabIndex = 12;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button5);
            this.groupBox2.Controls.Add(this.dataGridView1);
            this.groupBox2.Location = new System.Drawing.Point(763, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(220, 499);
            this.groupBox2.TabIndex = 24;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Ordenes Activas";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(6, 469);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(208, 24);
            this.button5.TabIndex = 24;
            this.button5.Text = "Dar Permisos";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2});
            this.dataGridView1.Location = new System.Drawing.Point(6, 19);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dataGridView1.Size = new System.Drawing.Size(208, 444);
            this.dataGridView1.TabIndex = 25;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Orden #";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Permiso";
            this.Column2.Name = "Column2";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(342, 61);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(28, 13);
            this.label5.TabIndex = 27;
            this.label5.Text = "Tipo";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Administrador",
            "Supervisor",
            "Almacenista",
            "Supervisor de Producción",
            "Supervisor de Cuchillas"});
            this.comboBox1.Location = new System.Drawing.Point(378, 58);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(234, 21);
            this.comboBox1.TabIndex = 26;
            // 
            // frmUsuarios
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(995, 523);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.gridSupervisor);
            this.Name = "frmUsuarios";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Usuarios";
            ((System.ComponentModel.ISupportInitialize)(this.gridSupervisor)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView gridSupervisor;
        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox txtEmpleado;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}