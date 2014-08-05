namespace Gimli
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBoxConsole = new System.Windows.Forms.TextBox();
            this.buttonSelectFile = new System.Windows.Forms.Button();
            this.buttonImportProfile = new System.Windows.Forms.Button();
            this.textBoxFile = new System.Windows.Forms.TextBox();
            this.buttonImportScores = new System.Windows.Forms.Button();
            this.generateReport = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBoxConsole
            // 
            this.textBoxConsole.BackColor = System.Drawing.Color.Black;
            this.textBoxConsole.Font = new System.Drawing.Font("Lucida Console", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxConsole.ForeColor = System.Drawing.Color.LawnGreen;
            this.textBoxConsole.Location = new System.Drawing.Point(85, 238);
            this.textBoxConsole.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textBoxConsole.Multiline = true;
            this.textBoxConsole.Name = "textBoxConsole";
            this.textBoxConsole.ReadOnly = true;
            this.textBoxConsole.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxConsole.Size = new System.Drawing.Size(943, 369);
            this.textBoxConsole.TabIndex = 11;
            this.textBoxConsole.Text = "> Started";
            // 
            // buttonSelectFile
            // 
            this.buttonSelectFile.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSelectFile.Location = new System.Drawing.Point(326, 30);
            this.buttonSelectFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.buttonSelectFile.Name = "buttonSelectFile";
            this.buttonSelectFile.Size = new System.Drawing.Size(74, 37);
            this.buttonSelectFile.TabIndex = 14;
            this.buttonSelectFile.Text = "File";
            this.buttonSelectFile.UseVisualStyleBackColor = true;
            this.buttonSelectFile.Click += new System.EventHandler(this.buttonSelectFile_Click);
            // 
            // buttonImportProfile
            // 
            this.buttonImportProfile.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonImportProfile.Location = new System.Drawing.Point(492, 94);
            this.buttonImportProfile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.buttonImportProfile.Name = "buttonImportProfile";
            this.buttonImportProfile.Size = new System.Drawing.Size(170, 58);
            this.buttonImportProfile.TabIndex = 15;
            this.buttonImportProfile.Text = "Import Profile";
            this.buttonImportProfile.UseVisualStyleBackColor = true;
            this.buttonImportProfile.Click += new System.EventHandler(this.buttonImportProfile_Click);
            // 
            // textBoxFile
            // 
            this.textBoxFile.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFile.Location = new System.Drawing.Point(406, 40);
            this.textBoxFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textBoxFile.Name = "textBoxFile";
            this.textBoxFile.ReadOnly = true;
            this.textBoxFile.Size = new System.Drawing.Size(602, 25);
            this.textBoxFile.TabIndex = 13;
            // 
            // buttonImportScores
            // 
            this.buttonImportScores.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonImportScores.Location = new System.Drawing.Point(814, 94);
            this.buttonImportScores.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.buttonImportScores.Name = "buttonImportScores";
            this.buttonImportScores.Size = new System.Drawing.Size(170, 58);
            this.buttonImportScores.TabIndex = 16;
            this.buttonImportScores.Text = "Import Scores";
            this.buttonImportScores.UseVisualStyleBackColor = true;
            this.buttonImportScores.Click += new System.EventHandler(this.buttonImportScores_Click);
            // 
            // generateReport
            // 
            this.generateReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateReport.Location = new System.Drawing.Point(353, 614);
            this.generateReport.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.generateReport.Name = "generateReport";
            this.generateReport.Size = new System.Drawing.Size(254, 43);
            this.generateReport.TabIndex = 17;
            this.generateReport.Text = "Generate Report";
            this.generateReport.UseVisualStyleBackColor = true;
            this.generateReport.Visible = false;
            this.generateReport.Click += new System.EventHandler(this.generateReport_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Gimli.Properties.Resources.Quest_Data_Tool_Logo;
            this.pictureBox1.Location = new System.Drawing.Point(24, 15);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(285, 200);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 18;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1065, 695);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.generateReport);
            this.Controls.Add(this.buttonSelectFile);
            this.Controls.Add(this.buttonImportProfile);
            this.Controls.Add(this.textBoxFile);
            this.Controls.Add(this.buttonImportScores);
            this.Controls.Add(this.textBoxConsole);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Form1";
            this.Text = "Quest Data Tool";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        
        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBoxConsole;
        private System.Windows.Forms.Button buttonSelectFile;
        private System.Windows.Forms.Button buttonImportProfile;
        private System.Windows.Forms.TextBox textBoxFile;
        private System.Windows.Forms.Button buttonImportScores;
        private System.Windows.Forms.Button generateReport;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

