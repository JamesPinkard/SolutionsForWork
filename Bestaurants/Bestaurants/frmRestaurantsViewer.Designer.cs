namespace ArcMapClassLibrary1
{
    partial class frmRestaurantsViewer
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
            this.cmbLayers = new System.Windows.Forms.ComboBox();
            this.Layers = new System.Windows.Forms.Label();
            this.btnPrintTDS = new System.Windows.Forms.Button();
            this.btnPrintWQ = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cmbLayers
            // 
            this.cmbLayers.FormattingEnabled = true;
            this.cmbLayers.Location = new System.Drawing.Point(54, 12);
            this.cmbLayers.Name = "cmbLayers";
            this.cmbLayers.Size = new System.Drawing.Size(201, 21);
            this.cmbLayers.TabIndex = 0;
            this.cmbLayers.SelectedIndexChanged += new System.EventHandler(this.cmbLayer_SelectedIndexChanged);
            // 
            // Layers
            // 
            this.Layers.AutoSize = true;
            this.Layers.Location = new System.Drawing.Point(13, 12);
            this.Layers.Name = "Layers";
            this.Layers.Size = new System.Drawing.Size(38, 13);
            this.Layers.TabIndex = 1;
            this.Layers.Text = "Layers";
            this.Layers.Click += new System.EventHandler(this.label1_Click);
            // 
            // btnPrintTDS
            // 
            this.btnPrintTDS.Location = new System.Drawing.Point(261, 10);
            this.btnPrintTDS.Name = "btnPrintTDS";
            this.btnPrintTDS.Size = new System.Drawing.Size(75, 23);
            this.btnPrintTDS.TabIndex = 2;
            this.btnPrintTDS.Text = "Print TDS";
            this.btnPrintTDS.UseVisualStyleBackColor = true;
            this.btnPrintTDS.Click += new System.EventHandler(this.btnPrintTDS_Click);
            // 
            // btnPrintWQ
            // 
            this.btnPrintWQ.Location = new System.Drawing.Point(342, 10);
            this.btnPrintWQ.Name = "btnPrintWQ";
            this.btnPrintWQ.Size = new System.Drawing.Size(75, 23);
            this.btnPrintWQ.TabIndex = 3;
            this.btnPrintWQ.Text = "Print WQ";
            this.btnPrintWQ.UseVisualStyleBackColor = true;
            this.btnPrintWQ.Click += new System.EventHandler(this.btnPrintWQ_Click);
            // 
            // frmRestaurantsViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(541, 101);
            this.Controls.Add(this.btnPrintWQ);
            this.Controls.Add(this.btnPrintTDS);
            this.Controls.Add(this.Layers);
            this.Controls.Add(this.cmbLayers);
            this.Name = "frmRestaurantsViewer";
            this.Text = "frmRestaurantsViewer";
            this.Load += new System.EventHandler(this.frmRestaurantsViewer_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ComboBox cmbLayers;
        private System.Windows.Forms.Label Layers;
        private System.Windows.Forms.Button btnPrintTDS;
        private System.Windows.Forms.Button btnPrintWQ;
    }
}