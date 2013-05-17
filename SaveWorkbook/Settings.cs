using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SaveWorkbook
{
    public partial class frmSettings : Form
    {
        public frmSettings()
        {
            InitializeComponent();
        }

        #region Buttons
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                txtPath.Text = path;
        }

        private void btnGapsBrowse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                txtPath.Text = path;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.PathSave = txtPath.Text;
            Properties.Settings.Default.Save();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        private bool SetPath(out string path)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            bool result;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                path = fd.SelectedPath;
                result = true;
            }
            else
            {
                path = String.Empty;
                result = false;
            }

            fd.Dispose();

            return result;
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            txtPath.Text = Properties.Settings.Default.PathSave;
        }
    }
}
