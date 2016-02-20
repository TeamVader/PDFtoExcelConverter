namespace PDFtoExcelConverter
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
            this.buttopen = new System.Windows.Forms.Button();
            this.buttconvert = new System.Windows.Forms.Button();
            this.labelname = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbocopies = new System.Windows.Forms.ComboBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.labelstopwatch = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // buttopen
            // 
            this.buttopen.Location = new System.Drawing.Point(12, 69);
            this.buttopen.Name = "buttopen";
            this.buttopen.Size = new System.Drawing.Size(108, 33);
            this.buttopen.TabIndex = 0;
            this.buttopen.Text = "open PDF ";
            this.buttopen.UseVisualStyleBackColor = true;
            this.buttopen.Click += new System.EventHandler(this.buttopen_Click);
            // 
            // buttconvert
            // 
            this.buttconvert.Location = new System.Drawing.Point(13, 109);
            this.buttconvert.Name = "buttconvert";
            this.buttconvert.Size = new System.Drawing.Size(107, 32);
            this.buttconvert.TabIndex = 1;
            this.buttconvert.Text = "convert";
            this.buttconvert.UseVisualStyleBackColor = true;
            this.buttconvert.Click += new System.EventHandler(this.buttconvert_Click);
            // 
            // labelname
            // 
            this.labelname.AutoSize = true;
            this.labelname.Location = new System.Drawing.Point(13, 13);
            this.labelname.Name = "labelname";
            this.labelname.Size = new System.Drawing.Size(0, 17);
            this.labelname.TabIndex = 2;
            // 
            // pictureBox1
            // 
            this.pictureBox1.ErrorImage = null;
            this.pictureBox1.Image = global::PDFtoExcelConverter.Properties.Resources.Stark_Industries_logo;
            this.pictureBox1.Location = new System.Drawing.Point(157, 86);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(476, 104);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 156);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "No of Copies";
            // 
            // cbocopies
            // 
            this.cbocopies.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbocopies.Location = new System.Drawing.Point(16, 177);
            this.cbocopies.Name = "cbocopies";
            this.cbocopies.Size = new System.Drawing.Size(104, 24);
            this.cbocopies.TabIndex = 5;
            this.cbocopies.SelectedIndexChanged += new System.EventHandler(this.cbocopies_SelectedIndexChanged);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(16, 40);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(421, 23);
            this.progressBar1.TabIndex = 6;
            this.progressBar1.Visible = false;
            // 
            // labelstopwatch
            // 
            this.labelstopwatch.AutoSize = true;
            this.labelstopwatch.Location = new System.Drawing.Point(504, 197);
            this.labelstopwatch.Name = "labelstopwatch";
            this.labelstopwatch.Size = new System.Drawing.Size(0, 17);
            this.labelstopwatch.TabIndex = 7;
            this.labelstopwatch.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(645, 218);
            this.Controls.Add(this.labelstopwatch);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.cbocopies);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.labelname);
            this.Controls.Add(this.buttconvert);
            this.Controls.Add(this.buttopen);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "PDF to Excel Converter";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttopen;
        private System.Windows.Forms.Button buttconvert;
        private System.Windows.Forms.Label labelname;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbocopies;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label labelstopwatch;
    }
}

