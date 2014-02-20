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
        Button btnGapsBrowse;
        TextBox txtGapsPath;
        Label lblGapsPath;

        Button btn117Browse;
        TextBox txt117Path;
        Label lbl117Path;

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
            //Gaps Button
            btnGapsBrowse = new Button();
            btnGapsBrowse.Text = "Browse";
            btnGapsBrowse.Name = "btnGapsBrowse";
            btnGapsBrowse.Location = new Point(329, 10);
            btnGapsBrowse.Click += btnGapsBrowse_Click;
            this.Controls.Add(btnGapsBrowse);

            //Gaps Textbox
            txtGapsPath = new TextBox();
            txtGapsPath.Width = 244;
            txtGapsPath.Height = 20;
            txtGapsPath.Name = "txtGapsPath";
            txtGapsPath.Leave += txtGapsPath_Leave;
            txtGapsPath.Location = new Point(btnGapsBrowse.Location.X - 250, btnGapsBrowse.Location.Y + 2);
            this.Controls.Add(txtGapsPath);

            //Gaps Label
            lblGapsPath = new Label();
            lblGapsPath.Size = new System.Drawing.Size(10, 10);
            lblGapsPath.Text = "GAPS Path";
            lblGapsPath.Name = "lblGapsPath";
            this.Controls.Add(lblGapsPath);
            lblGapsPath.AutoSize = true;
            lblGapsPath.Location = new Point(txtGapsPath.Location.X - lblGapsPath.Width - 7, txtGapsPath.Location.Y + 3);

            //117
            btn117Browse = new Button();
            btn117Browse.Text = "Browse";
            btn117Browse.Name = "btn117Browse";
            btn117Browse.Location = new Point(329, 36); // Y = previous buttion.Y + 26
            this.Controls.Add(btn117Browse);

            txt117Path = new TextBox();
            txt117Path.Width = 244;
            txt117Path.Height = 20;
            txt117Path.Name = "txt117Path";
            txt117Path.Location = new Point(btn117Browse.Location.X - 250, btn117Browse.Location.Y + 2);
            this.Controls.Add(txt117Path);
            
            lbl117Path = new Label();

            //Text
            txt117Path.Text = Properties.Settings.Default.Path117;
            txt473Path.Text = Properties.Settings.Default.Path473;
            txtGapsPath.Text = Properties.Settings.Default.PathGAPS;
            txt325Path.Text = Properties.Settings.Default.Path325;
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

        static private int StringWidth(Graphics graphics, string text, Font font)
        {
            System.Drawing.StringFormat format = new System.Drawing.StringFormat();
            System.Drawing.RectangleF rect = new System.Drawing.RectangleF(0, 0,
                                                                          1000, 1000);
            System.Drawing.CharacterRange[] ranges = { new System.Drawing.CharacterRange(0, text.Length) };
            System.Drawing.Region[] regions = new System.Drawing.Region[1];

            format.SetMeasurableCharacterRanges(ranges);

            regions = graphics.MeasureCharacterRanges(text, font, rect, format);
            rect = regions[0].GetBounds(graphics);

            return (int)(rect.Right + 1.0f);
        }
    }
}
