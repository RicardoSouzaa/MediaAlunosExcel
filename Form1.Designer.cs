
namespace MediaAlunosExcel
{
    partial class Form1
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
                SalvarPlanilha();
                excelApp.Quit();
                this.Close();
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
            this.lblNome = new System.Windows.Forms.Label();
            this.lblNota1 = new System.Windows.Forms.Label();
            this.lblMedia = new System.Windows.Forms.Label();
            this.lblNota2 = new System.Windows.Forms.Label();
            this.txtNome = new System.Windows.Forms.TextBox();
            this.txtMedia = new System.Windows.Forms.TextBox();
            this.cbNota1 = new System.Windows.Forms.ComboBox();
            this.cbNota2 = new System.Windows.Forms.ComboBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnImprimir = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblNome
            // 
            this.lblNome.AutoSize = true;
            this.lblNome.Location = new System.Drawing.Point(13, 13);
            this.lblNome.Name = "lblNome";
            this.lblNome.Size = new System.Drawing.Size(47, 15);
            this.lblNome.TabIndex = 0;
            this.lblNome.Text = "NOME:";
            // 
            // lblNota1
            // 
            this.lblNota1.AutoSize = true;
            this.lblNota1.Location = new System.Drawing.Point(13, 70);
            this.lblNota1.Name = "lblNota1";
            this.lblNota1.Size = new System.Drawing.Size(46, 15);
            this.lblNota1.TabIndex = 1;
            this.lblNota1.Text = "NOTA1";
            // 
            // lblMedia
            // 
            this.lblMedia.AutoSize = true;
            this.lblMedia.Location = new System.Drawing.Point(13, 127);
            this.lblMedia.Name = "lblMedia";
            this.lblMedia.Size = new System.Drawing.Size(45, 15);
            this.lblMedia.TabIndex = 2;
            this.lblMedia.Text = "MÉDIA";
            // 
            // lblNota2
            // 
            this.lblNota2.AutoSize = true;
            this.lblNota2.Location = new System.Drawing.Point(205, 70);
            this.lblNota2.Name = "lblNota2";
            this.lblNota2.Size = new System.Drawing.Size(46, 15);
            this.lblNota2.TabIndex = 1;
            this.lblNota2.Text = "NOTA2";
            // 
            // txtNome
            // 
            this.txtNome.Location = new System.Drawing.Point(67, 7);
            this.txtNome.Name = "txtNome";
            this.txtNome.Size = new System.Drawing.Size(309, 20);
            this.txtNome.TabIndex = 3;
            // 
            // txtMedia
            // 
            this.txtMedia.Location = new System.Drawing.Point(67, 122);
            this.txtMedia.Name = "txtMedia";
            this.txtMedia.Size = new System.Drawing.Size(110, 20);
            this.txtMedia.TabIndex = 4;
            // 
            // cbNota1
            // 
            this.cbNota1.FormattingEnabled = true;
            this.cbNota1.Location = new System.Drawing.Point(67, 67);
            this.cbNota1.Name = "cbNota1";
            this.cbNota1.Size = new System.Drawing.Size(110, 21);
            this.cbNota1.TabIndex = 5;
            // 
            // cbNota2
            // 
            this.cbNota2.FormattingEnabled = true;
            this.cbNota2.Location = new System.Drawing.Point(266, 67);
            this.cbNota2.Name = "cbNota2";
            this.cbNota2.Size = new System.Drawing.Size(110, 21);
            this.cbNota2.TabIndex = 6;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(16, 151);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(161, 23);
            this.btnOk.TabIndex = 7;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnImprimir
            // 
            this.btnImprimir.Location = new System.Drawing.Point(208, 151);
            this.btnImprimir.Name = "btnImprimir";
            this.btnImprimir.Size = new System.Drawing.Size(168, 23);
            this.btnImprimir.TabIndex = 8;
            this.btnImprimir.Text = "IMPRIMIR";
            this.btnImprimir.UseVisualStyleBackColor = true;
            this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(401, 208);
            this.Controls.Add(this.btnImprimir);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.cbNota2);
            this.Controls.Add(this.cbNota1);
            this.Controls.Add(this.txtMedia);
            this.Controls.Add(this.txtNome);
            this.Controls.Add(this.lblMedia);
            this.Controls.Add(this.lblNota2);
            this.Controls.Add(this.lblNota1);
            this.Controls.Add(this.lblNome);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Média Excel";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblNome;
        private System.Windows.Forms.Label lblNota1;
        private System.Windows.Forms.Label lblMedia;
        private System.Windows.Forms.Label lblNota2;
        private System.Windows.Forms.TextBox txtNome;
        private System.Windows.Forms.TextBox txtMedia;
        private System.Windows.Forms.ComboBox cbNota1;
        private System.Windows.Forms.ComboBox cbNota2;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnImprimir;
    }
}

