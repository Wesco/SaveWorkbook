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
                txtGapsPath.Text = path;
        }

        private void btn117Browse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                txt117Path.Text = path;
        }

        private void btn473Browse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                txt473Path.Text = path;
        }

        private void btn325Browse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                txt325Path.Text = path;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {

            if (IsValidPath(txt117Path.Text) & IsValidPath(txt473Path.Text) & IsValidPath(txtGapsPath.Text))
            {
                if (IsValidPath(txt117Path.Text))
                    Properties.Settings.Default.Path117 = txt117Path.Text;

                if (IsValidPath(txt473Path.Text))
                    Properties.Settings.Default.Path473 = txt473Path.Text;

                if (IsValidPath(txtGapsPath.Text))
                    Properties.Settings.Default.PathGAPS = txtGapsPath.Text;

                if (IsValidPath(txt325Path.Text))
                    Properties.Settings.Default.Path325 = txt325Path.Text;

                Properties.Settings.Default.Save();
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region LostFocus events
        private void txtGapsPath_Leave(object sender, EventArgs e)
        {
            if (txtGapsPath.Text.Right(1) != @"\")
                txtGapsPath.Text += @"\";

            if (IsValidPath(txtGapsPath.Text))
            {
                txtGapsPath.BackColor = Color.White;
                Properties.Settings.Default.PathGAPS = txtGapsPath.Text;
            }
            else
                txtGapsPath.BackColor = Color.LightPink;
        }

        private void txt117Path_Leave(object sender, EventArgs e)
        {
            if (txt117Path.Text.Right(1) != @"\")
                txt117Path.Text += @"\";

            if (IsValidPath(txt117Path.Text))
            {
                txt117Path.BackColor = Color.White;
                Properties.Settings.Default.Path117 = txt117Path.Text;
            }
            else
                txt117Path.BackColor = Color.LightPink;
        }

        private void txt473Path_Leave(object sender, EventArgs e)
        {
            if (txt473Path.Text.Right(1) != @"\")
                txt473Path.Text += @"\";

            if (IsValidPath(txt473Path.Text))
            {
                txt473Path.BackColor = Color.White;
                Properties.Settings.Default.Path473 = txt473Path.Text;
            }
            else
                txt473Path.BackColor = Color.LightPink;
        }

        private void txt325Path_Leave(object sender, EventArgs e)
        {
            if (txt325Path.Text.Right(1) != @"\")
                txt325Path.Text += @"\";

            if (IsValidPath(txt325Path.Text))
            {
                txt325Path.BackColor = Color.White;
                Properties.Settings.Default.Path325 = txt473Path.Text;
            }
            else
                txt473Path.BackColor = Color.LightPink;
        }
        #endregion

        private void frmSettings_Load(object sender, EventArgs e)
        {
            txt117Path.Text = Properties.Settings.Default.Path117;
            txt473Path.Text = Properties.Settings.Default.Path473;
            txtGapsPath.Text = Properties.Settings.Default.PathGAPS;
        }

        private bool SetPath(out string path)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            bool result;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                path = fd.SelectedPath;
                if (path.Substring(path.Length - 1) != "\\")
                    path += "\\";

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

        private bool IsValidPath(string path)
        {
            Regex rxFilePath = new Regex(@"^(?:[A-Za-z]\:\\|\\\\[\w.]+\\)(?:[^\\ ][\w!@#$%^&()_+;'\.,  .]*\\)*$");
            return rxFilePath.IsMatch(path);
        }
    }
}
