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

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string path = txtPath.Text;

            if (SetPath(ref path))
            {
                txtPath.Text = path;
                Properties.Settings.Default.PathSave = path;
                Properties.Settings.Default.Save();
            }

            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private bool SetPath(ref string path)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            bool result;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                path = fd.SelectedPath;
                result = true;
            }
            else
                result = false;
   
            fd.Dispose();

            return result;
        }
    }
}
