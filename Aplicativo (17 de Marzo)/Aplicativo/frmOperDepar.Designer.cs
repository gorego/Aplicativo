namespace Aplicativo
{
    partial class frmOperDepar
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
            this.btnCerrar = new System.Windows.Forms.Button();
            this.gridOperadores = new System.Windows.Forms.DataGridView();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.gridOperadores)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCerrar
            // 
            this.btnCerrar.Location = new System.Drawing.Point(12, 255);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(916, 33);
            this.btnCerrar.TabIndex = 1;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // gridOperadores
            // 
            this.gridOperadores.AllowUserToAddRows = false;
            this.gridOperadores.AllowUserToDeleteRows = false;
            this.gridOperadores.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridOperadores.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.gridOperadores.BackgroundColor = System.Drawing.SystemColors.Window;
            this.gridOperadores.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.gridOperadores.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.gridOperadores.ColumnHeadersHeight = 24;
            this.gridOperadores.GridColor = System.Drawing.SystemColors.ButtonFace;
            this.gridOperadores.Location = new System.Drawing.Point(12, 27);
            this.gridOperadores.Name = "gridOperadores";
            this.gridOperadores.ReadOnly = true;
            this.gridOperadores.RowHeadersVisible = false;
            this.gridOperadores.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.gridOperadores.Size = new System.Drawing.Size(916, 222);
            this.gridOperadores.TabIndex = 2;
            this.gridOperadores.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridOperadores_CellClick);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(831, 11);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(84, 13);
            this.linkLabel1.TabIndex = 76;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Exportar a Excel";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // frmOperDepar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(940, 300);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.gridOperadores);
            this.Controls.Add(this.btnCerrar);
            this.Name = "frmOperDepar";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Operadores";
            ((System.ComponentModel.ISupportInitialize)(this.gridOperadores)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCerrar;
        private System.Windows.Forms.DataGridView gridOperadores;
        private System.Windows.Forms.LinkLabel linkLabel1;
    }
}