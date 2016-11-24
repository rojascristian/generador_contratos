namespace AplicacionGeneradorContratos
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnGenerar = new System.Windows.Forms.Button();
            this.lblPathExcel = new System.Windows.Forms.Label();
            this.tbPathExcel = new System.Windows.Forms.TextBox();
            this.btnImportarExcel = new System.Windows.Forms.Button();
            this.lblDestino = new System.Windows.Forms.Label();
            this.tbDestino = new System.Windows.Forms.TextBox();
            this.btnSeleccionarCarpeta = new System.Windows.Forms.Button();
            this.btnDescargarExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnGenerar
            // 
            this.btnGenerar.Enabled = false;
            this.btnGenerar.Location = new System.Drawing.Point(439, 115);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(75, 23);
            this.btnGenerar.TabIndex = 2;
            this.btnGenerar.Text = "Generar";
            this.btnGenerar.UseVisualStyleBackColor = true;
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // lblPathExcel
            // 
            this.lblPathExcel.AutoSize = true;
            this.lblPathExcel.Location = new System.Drawing.Point(8, 17);
            this.lblPathExcel.Name = "lblPathExcel";
            this.lblPathExcel.Size = new System.Drawing.Size(91, 13);
            this.lblPathExcel.TabIndex = 5;
            this.lblPathExcel.Text = "Ruta fuente excel";
            this.lblPathExcel.Click += new System.EventHandler(this.label1_Click);
            // 
            // tbPathExcel
            // 
            this.tbPathExcel.Enabled = false;
            this.tbPathExcel.Location = new System.Drawing.Point(105, 14);
            this.tbPathExcel.Name = "tbPathExcel";
            this.tbPathExcel.Size = new System.Drawing.Size(266, 20);
            this.tbPathExcel.TabIndex = 6;
            this.tbPathExcel.TextChanged += new System.EventHandler(this.tbPathExcel_TextChanged_1);
            // 
            // btnImportarExcel
            // 
            this.btnImportarExcel.Location = new System.Drawing.Point(381, 12);
            this.btnImportarExcel.Name = "btnImportarExcel";
            this.btnImportarExcel.Size = new System.Drawing.Size(133, 23);
            this.btnImportarExcel.TabIndex = 7;
            this.btnImportarExcel.Text = "Seleccionar Archivo";
            this.btnImportarExcel.UseVisualStyleBackColor = true;
            this.btnImportarExcel.Click += new System.EventHandler(this.btnImportarExcel_Click);
            // 
            // lblDestino
            // 
            this.lblDestino.AutoSize = true;
            this.lblDestino.Location = new System.Drawing.Point(8, 72);
            this.lblDestino.Name = "lblDestino";
            this.lblDestino.Size = new System.Drawing.Size(81, 13);
            this.lblDestino.TabIndex = 8;
            this.lblDestino.Text = "Carpeta destino";
            // 
            // tbDestino
            // 
            this.tbDestino.Enabled = false;
            this.tbDestino.Location = new System.Drawing.Point(105, 72);
            this.tbDestino.Name = "tbDestino";
            this.tbDestino.Size = new System.Drawing.Size(266, 20);
            this.tbDestino.TabIndex = 9;
            this.tbDestino.TextChanged += new System.EventHandler(this.tbDestino_TextChanged_1);
            // 
            // btnSeleccionarCarpeta
            // 
            this.btnSeleccionarCarpeta.Location = new System.Drawing.Point(381, 70);
            this.btnSeleccionarCarpeta.Name = "btnSeleccionarCarpeta";
            this.btnSeleccionarCarpeta.Size = new System.Drawing.Size(133, 23);
            this.btnSeleccionarCarpeta.TabIndex = 10;
            this.btnSeleccionarCarpeta.Text = "Seleccionar carpeta";
            this.btnSeleccionarCarpeta.UseVisualStyleBackColor = true;
            this.btnSeleccionarCarpeta.Click += new System.EventHandler(this.btnSeleccionarCarpeta_Click);
            // 
            // btnDescargarExcel
            // 
            this.btnDescargarExcel.Location = new System.Drawing.Point(381, 41);
            this.btnDescargarExcel.Name = "btnDescargarExcel";
            this.btnDescargarExcel.Size = new System.Drawing.Size(133, 23);
            this.btnDescargarExcel.TabIndex = 11;
            this.btnDescargarExcel.Text = "Descargar Plantilla Excel";
            this.btnDescargarExcel.UseVisualStyleBackColor = true;
            this.btnDescargarExcel.Click += new System.EventHandler(this.btnDescargarExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 153);
            this.Controls.Add(this.btnDescargarExcel);
            this.Controls.Add(this.btnSeleccionarCarpeta);
            this.Controls.Add(this.tbDestino);
            this.Controls.Add(this.lblDestino);
            this.Controls.Add(this.btnImportarExcel);
            this.Controls.Add(this.tbPathExcel);
            this.Controls.Add(this.lblPathExcel);
            this.Controls.Add(this.btnGenerar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generar Contratos";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.Label lblPathExcel;
        private System.Windows.Forms.TextBox tbPathExcel;
        private System.Windows.Forms.Button btnImportarExcel;
        private System.Windows.Forms.Label lblDestino;
        private System.Windows.Forms.TextBox tbDestino;
        private System.Windows.Forms.Button btnSeleccionarCarpeta;
        private System.Windows.Forms.Button btnDescargarExcel;
    }
}

