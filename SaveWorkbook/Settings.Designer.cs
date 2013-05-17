namespace SaveWorkbook
{
    partial class frmSettings
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
            this.lblGapsPath = new System.Windows.Forms.Label();
            this.lbl117Path = new System.Windows.Forms.Label();
            this.lbl473Path = new System.Windows.Forms.Label();
            this.txt473Path = new System.Windows.Forms.TextBox();
            this.txt117Path = new System.Windows.Forms.TextBox();
            this.txtGapsPath = new System.Windows.Forms.TextBox();
            this.btnGapsBrowse = new System.Windows.Forms.Button();
            this.btn117Browse = new System.Windows.Forms.Button();
            this.btn473Browse = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblGapsPath
            // 
            this.lblGapsPath.AutoSize = true;
            this.lblGapsPath.Location = new System.Drawing.Point(8, 15);
            this.lblGapsPath.Name = "lblGapsPath";
            this.lblGapsPath.Size = new System.Drawing.Size(61, 13);
            this.lblGapsPath.TabIndex = 2;
            this.lblGapsPath.Text = "GAPS Path";
            this.lblGapsPath.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lbl117Path
            // 
            this.lbl117Path.AutoSize = true;
            this.lbl117Path.Location = new System.Drawing.Point(20, 41);
            this.lbl117Path.Name = "lbl117Path";
            this.lbl117Path.Size = new System.Drawing.Size(50, 13);
            this.lbl117Path.TabIndex = 3;
            this.lbl117Path.Text = "117 Path";
            this.lbl117Path.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lbl473Path
            // 
            this.lbl473Path.AutoSize = true;
            this.lbl473Path.Location = new System.Drawing.Point(20, 67);
            this.lbl473Path.Name = "lbl473Path";
            this.lbl473Path.Size = new System.Drawing.Size(50, 13);
            this.lbl473Path.TabIndex = 4;
            this.lbl473Path.Text = "473 Path";
            this.lbl473Path.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txt473Path
            // 
            this.txt473Path.Location = new System.Drawing.Point(79, 64);
            this.txt473Path.Name = "txt473Path";
            this.txt473Path.Size = new System.Drawing.Size(244, 20);
            this.txt473Path.TabIndex = 5;
            // 
            // txt117Path
            // 
            this.txt117Path.Location = new System.Drawing.Point(79, 38);
            this.txt117Path.Name = "txt117Path";
            this.txt117Path.Size = new System.Drawing.Size(244, 20);
            this.txt117Path.TabIndex = 6;
            // 
            // txtGapsPath
            // 
            this.txtGapsPath.Location = new System.Drawing.Point(79, 12);
            this.txtGapsPath.Name = "txtGapsPath";
            this.txtGapsPath.Size = new System.Drawing.Size(244, 20);
            this.txtGapsPath.TabIndex = 7;
            // 
            // btnGapsBrowse
            // 
            this.btnGapsBrowse.Location = new System.Drawing.Point(329, 10);
            this.btnGapsBrowse.Name = "btnGapsBrowse";
            this.btnGapsBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnGapsBrowse.TabIndex = 9;
            this.btnGapsBrowse.Text = "Browse";
            this.btnGapsBrowse.UseVisualStyleBackColor = true;
            this.btnGapsBrowse.Click += new System.EventHandler(this.btnGapsBrowse_Click);
            // 
            // btn117Browse
            // 
            this.btn117Browse.Location = new System.Drawing.Point(329, 36);
            this.btn117Browse.Name = "btn117Browse";
            this.btn117Browse.Size = new System.Drawing.Size(75, 23);
            this.btn117Browse.TabIndex = 10;
            this.btn117Browse.Text = "Browse";
            this.btn117Browse.UseVisualStyleBackColor = true;
            this.btn117Browse.Click += new System.EventHandler(this.btn117Browse_Click);
            // 
            // btn473Browse
            // 
            this.btn473Browse.Location = new System.Drawing.Point(329, 62);
            this.btn473Browse.Name = "btn473Browse";
            this.btn473Browse.Size = new System.Drawing.Size(75, 23);
            this.btn473Browse.TabIndex = 11;
            this.btn473Browse.Text = "Browse";
            this.btn473Browse.UseVisualStyleBackColor = true;
            this.btn473Browse.Click += new System.EventHandler(this.btn473Browse_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(329, 104);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(248, 104);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 13;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(416, 139);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btn473Browse);
            this.Controls.Add(this.btn117Browse);
            this.Controls.Add(this.btnGapsBrowse);
            this.Controls.Add(this.txtGapsPath);
            this.Controls.Add(this.txt117Path);
            this.Controls.Add(this.txt473Path);
            this.Controls.Add(this.lbl473Path);
            this.Controls.Add(this.lbl117Path);
            this.Controls.Add(this.lblGapsPath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmSettings";
            this.Text = "Settings";
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblGapsPath;
        private System.Windows.Forms.Label lbl117Path;
        private System.Windows.Forms.Label lbl473Path;
        private System.Windows.Forms.TextBox txt473Path;
        private System.Windows.Forms.TextBox txt117Path;
        private System.Windows.Forms.TextBox txtGapsPath;
        private System.Windows.Forms.Button btnGapsBrowse;
        private System.Windows.Forms.Button btn117Browse;
        private System.Windows.Forms.Button btn473Browse;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
    }
}