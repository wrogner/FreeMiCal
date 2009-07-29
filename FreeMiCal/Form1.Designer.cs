namespace FreeMiCal
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
		protected override void Dispose (bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose ();
			}
			base.Dispose (disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent ()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager (typeof (Form1));
			this.pictureBox1 = new System.Windows.Forms.PictureBox ();
			this.label2 = new System.Windows.Forms.Label ();
			this.label3 = new System.Windows.Forms.Label ();
			this.txtProfile = new System.Windows.Forms.TextBox ();
			this.label4 = new System.Windows.Forms.Label ();
			this.label5 = new System.Windows.Forms.Label ();
			this.numStart = new System.Windows.Forms.NumericUpDown ();
			this.numEnd = new System.Windows.Forms.NumericUpDown ();
			this.progressBar1 = new System.Windows.Forms.ProgressBar ();
			this.btnExport = new System.Windows.Forms.Button ();
			this.btnExit = new System.Windows.Forms.Button ();
			this.label6 = new System.Windows.Forms.Label ();
			this.txtFile = new System.Windows.Forms.TextBox ();
			this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker ();
			this.lblError = new System.Windows.Forms.Label ();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit ();
			((System.ComponentModel.ISupportInitialize)(this.numStart)).BeginInit ();
			((System.ComponentModel.ISupportInitialize)(this.numEnd)).BeginInit ();
			this.SuspendLayout ();
			// 
			// pictureBox1
			// 
			this.pictureBox1.BackColor = System.Drawing.SystemColors.Control;
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject ("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point (-2, -2);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size (627, 87);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 0;
			this.pictureBox1.TabStop = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font ("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.label2.Location = new System.Drawing.Point (40, 100);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size (538, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Free your Outlook calendar items by bulk exporting them to standard RFC 2445 iCal" +
				" format";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.label3.Location = new System.Drawing.Point (39, 151);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size (198, 24);
			this.label3.TabIndex = 3;
			this.label3.Text = "Default Outlook profile:";
			// 
			// txtProfile
			// 
			this.txtProfile.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtProfile.Location = new System.Drawing.Point (288, 148);
			this.txtProfile.Name = "txtProfile";
			this.txtProfile.Size = new System.Drawing.Size (172, 29);
			this.txtProfile.TabIndex = 6;
			this.txtProfile.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label4.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.label4.Location = new System.Drawing.Point (39, 205);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size (243, 24);
			this.label4.TabIndex = 5;
			this.label4.Text = "Export from Calendar item #";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label5.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.label5.Location = new System.Drawing.Point (393, 205);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size (80, 24);
			this.label5.TabIndex = 7;
			this.label5.Text = "to item #";
			// 
			// numStart
			// 
			this.numStart.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.numStart.Location = new System.Drawing.Point (288, 203);
			this.numStart.Minimum = new decimal (new int[] {
            1,
            0,
            0,
            0});
			this.numStart.Name = "numStart";
			this.numStart.Size = new System.Drawing.Size (99, 29);
			this.numStart.TabIndex = 2;
			this.numStart.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.numStart.ThousandsSeparator = true;
			this.numStart.Value = new decimal (new int[] {
            1,
            0,
            0,
            0});
			// 
			// numEnd
			// 
			this.numEnd.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.numEnd.Location = new System.Drawing.Point (479, 203);
			this.numEnd.Minimum = new decimal (new int[] {
            1,
            0,
            0,
            0});
			this.numEnd.Name = "numEnd";
			this.numEnd.Size = new System.Drawing.Size (99, 29);
			this.numEnd.TabIndex = 3;
			this.numEnd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.numEnd.ThousandsSeparator = true;
			this.numEnd.Value = new decimal (new int[] {
            1,
            0,
            0,
            0});
			// 
			// progressBar1
			// 
			this.progressBar1.ForeColor = System.Drawing.SystemColors.HotTrack;
			this.progressBar1.Location = new System.Drawing.Point (43, 350);
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size (535, 18);
			this.progressBar1.TabIndex = 11;
			this.progressBar1.Visible = false;
			// 
			// btnExport
			// 
			this.btnExport.Font = new System.Drawing.Font ("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnExport.Location = new System.Drawing.Point (406, 393);
			this.btnExport.Name = "btnExport";
			this.btnExport.Size = new System.Drawing.Size (172, 35);
			this.btnExport.TabIndex = 1;
			this.btnExport.Text = "Fr&ee them...";
			this.btnExport.UseVisualStyleBackColor = true;
			this.btnExport.Click += new System.EventHandler (this.btnExport_Click);
			// 
			// btnExit
			// 
			this.btnExit.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnExit.Location = new System.Drawing.Point (43, 393);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size (83, 35);
			this.btnExit.TabIndex = 5;
			this.btnExit.Text = "E&xit";
			this.btnExit.UseVisualStyleBackColor = true;
			this.btnExit.Click += new System.EventHandler (this.btnExit_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label6.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.label6.Location = new System.Drawing.Point (39, 262);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size (118, 24);
			this.label6.TabIndex = 12;
			this.label6.Text = "Export to file:";
			// 
			// txtFile
			// 
			this.txtFile.Font = new System.Drawing.Font ("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtFile.Location = new System.Drawing.Point (289, 259);
			this.txtFile.Name = "txtFile";
			this.txtFile.Size = new System.Drawing.Size (289, 29);
			this.txtFile.TabIndex = 4;
			// 
			// backgroundWorker1
			// 
			this.backgroundWorker1.WorkerReportsProgress = true;
			this.backgroundWorker1.WorkerSupportsCancellation = true;
			this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler (this.backgroundWorker1_DoWork);
			this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler (this.backgroundWorker1_RunWorkerCompleted);
			this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler (this.backgroundWorker1_ProgressChanged);
			// 
			// lblError
			// 
			this.lblError.Font = new System.Drawing.Font ("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblError.Location = new System.Drawing.Point (40, 311);
			this.lblError.Name = "lblError";
			this.lblError.Size = new System.Drawing.Size (538, 31);
			this.lblError.TabIndex = 13;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF (6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.Window;
			this.ClientSize = new System.Drawing.Size (623, 466);
			this.Controls.Add (this.lblError);
			this.Controls.Add (this.txtFile);
			this.Controls.Add (this.label6);
			this.Controls.Add (this.btnExit);
			this.Controls.Add (this.btnExport);
			this.Controls.Add (this.progressBar1);
			this.Controls.Add (this.numEnd);
			this.Controls.Add (this.numStart);
			this.Controls.Add (this.label5);
			this.Controls.Add (this.label4);
			this.Controls.Add (this.txtProfile);
			this.Controls.Add (this.label3);
			this.Controls.Add (this.label2);
			this.Controls.Add (this.pictureBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "Form1";
			this.Text = "FreeMiCal";
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit ();
			((System.ComponentModel.ISupportInitialize)(this.numStart)).EndInit ();
			((System.ComponentModel.ISupportInitialize)(this.numEnd)).EndInit ();
			this.ResumeLayout (false);
			this.PerformLayout ();

		}

		#endregion

		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtProfile;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.NumericUpDown numStart;
		private System.Windows.Forms.NumericUpDown numEnd;
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.Button btnExport;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtFile;
		private System.ComponentModel.BackgroundWorker backgroundWorker1;
		private System.Windows.Forms.Label lblError;
	}
}