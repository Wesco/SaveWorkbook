using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private void btnGapsBrowse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                Properties.Settings.Default.PathGAPS = path;
        }

        private void btn117Browse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                Properties.Settings.Default.Path117 = path;
        }

        private void btn473Browse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                Properties.Settings.Default.Path473 = path;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Regex rxFilePath = new Regex(@"^(?:[A-Za-z]\:\\|\\\\[\w.]+\\)(?:[^\\ ][\w!@#$%^&()_+;'\.,  .]*\\)+$");

            if (rxFilePath.IsMatch(txt117Path.Text))
            {
                txt117Path.BackColor = Color.White;
                Properties.Settings.Default.Path117 = txt117Path.Text;
            }
            else
                txt117Path.BackColor = Color.LightPink;
                

            if (rxFilePath.IsMatch(txt473Path.Text))
                Properties.Settings.Default.Path473 = txt473Path.Text;

            if (rxFilePath.IsMatch(txtGapsPath.Text))
                Properties.Settings.Default.PathGAPS = txtGapsPath.Text;

            if (rxFilePath.IsMatch(txt117Path.Text) &
                rxFilePath.IsMatch(txt473Path.Text) &
                rxFilePath.IsMatch(txtGapsPath.Text))
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
            txt117Path.Text = Properties.Settings.Default.Path117;
            txt473Path.Text = Properties.Settings.Default.Path473;
            txtGapsPath.Text = Properties.Settings.Default.PathGAPS;
        }
    }
}
