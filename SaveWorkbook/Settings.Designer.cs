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
            this.lbl473Path = new System.Windows.Forms.Label();
            this.txt473Path = new System.Windows.Forms.TextBox();
            this.btn473Browse = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btn325Browse = new System.Windows.Forms.Button();
            this.txt325Path = new System.Windows.Forms.TextBox();
            this.lbl325Path = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lbl473Path
            // 
            this.lbl473Path.AutoSize = true;
            this.lbl473Path.Location = new System.Drawing.Point(20, 67);
            this.lbl473Path.Name = "lbl473Path";
            this.lbl473Path.Size = new System.Drawing.Size(50, 13);
            this.lbl473Path.TabIndex = 101;
            this.lbl473Path.Text = "473 Path";
            this.lbl473Path.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txt473Path
            // 
            this.txt473Path.Location = new System.Drawing.Point(79, 64);
            this.txt473Path.Name = "txt473Path";
            this.txt473Path.Size = new System.Drawing.Size(244, 20);
            this.txt473Path.TabIndex = 5;
            this.txt473Path.Leave += new System.EventHandler(this.txt473Path_Leave);
            // 
            // btn473Browse
            // 
            this.btn473Browse.Location = new System.Drawing.Point(329, 62);
            this.btn473Browse.Name = "btn473Browse";
            this.btn473Browse.Size = new System.Drawing.Size(75, 23);
            this.btn473Browse.TabIndex = 6;
            this.btn473Browse.Text = "Browse";
            this.btn473Browse.UseVisualStyleBackColor = true;
            this.btn473Browse.Click += new System.EventHandler(this.btn473Browse_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(329, 126);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(248, 126);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btn325Browse
            // 
            this.btn325Browse.Location = new System.Drawing.Point(329, 88);
            this.btn325Browse.Name = "btn325Browse";
            this.btn325Browse.Size = new System.Drawing.Size(75, 23);
            this.btn325Browse.TabIndex = 102;
            this.btn325Browse.Text = "Browse";
            this.btn325Browse.UseVisualStyleBackColor = true;
            this.btn325Browse.Click += new System.EventHandler(this.btn325Browse_Click);
            // 
            // txt325Path
            // 
            this.txt325Path.Location = new System.Drawing.Point(79, 91);
            this.txt325Path.Name = "txt325Path";
            this.txt325Path.Size = new System.Drawing.Size(244, 20);
            this.txt325Path.TabIndex = 103;
            this.txt325Path.Leave += new System.EventHandler(this.txt325Path_Leave);
            // 
            // lbl325Path
            // 
            this.lbl325Path.AutoSize = true;
            this.lbl325Path.Location = new System.Drawing.Point(19, 94);
            this.lbl325Path.Name = "lbl325Path";
            this.lbl325Path.Size = new System.Drawing.Size(50, 13);
            this.lbl325Path.TabIndex = 104;
            this.lbl325Path.Text = "325 Path";
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(416, 161);
            this.Controls.Add(this.lbl325Path);
            this.Controls.Add(this.txt325Path);
            this.Controls.Add(this.btn325Browse);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btn473Browse);
            this.Controls.Add(this.txt473Path);
            this.Controls.Add(this.lbl473Path);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "frmSettings";
            this.Text = "Settings";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl473Path;
        private System.Windows.Forms.TextBox txt473Path;
        private System.Windows.Forms.Button btn473Browse;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btn325Browse;
        private System.Windows.Forms.TextBox txt325Path;
        private System.Windows.Forms.Label lbl325Path;
    }
}