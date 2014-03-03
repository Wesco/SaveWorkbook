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
        #region Controls
        //Gaps
        Button btnGapsBrowse;
        TextBox txtGapsPath;
        Label lblGapsPath;

        //117
        Button btn117Browse;
        TextBox txt117Path;
        Label lbl117Path;

        //473
        Button btn473Browse;
        TextBox txt473Path;
        Label lbl473Path;

        //325
        Button btn325Browse;
        TextBox txt325Path;
        Label lbl325Path;

        //AP1000
        Button btnAP1000Browse;
        TextBox txtAP1000Path;
        Label lblAP1000Path;

        //Save
        Button btnSave;

        //Cancel
        Button btnCancel;
        #endregion

        public frmSettings()
        {
            InitializeComponent();
        }

        #region Button_Events
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

        void btnAP1000Browse_Click(object sender, EventArgs e)
        {
            string path;

            if (SetPath(out path))
                txtAP1000Path.Text = path;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    if (!IsValidPath(control.Text))
                        return;
                }
            }

            Properties.Settings.Default.Path117 = txt117Path.Text;
            Properties.Settings.Default.Path473 = txt473Path.Text;
            Properties.Settings.Default.PathGAPS = txtGapsPath.Text;
            Properties.Settings.Default.Path325 = txt325Path.Text;
            Properties.Settings.Default.PathAP1000 = txtAP1000Path.Text;

            Properties.Settings.Default.Save();
            MessageBox.Show("Settings Saved!", "Success");
            this.Close();
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
            int prevBtnY = 10;
            int txtCounter = 0;

            #region Gaps
            //Gaps Button
            btnGapsBrowse = new Button();
            btnGapsBrowse.Text = "Browse";
            btnGapsBrowse.Name = "btnGapsBrowse";
            btnGapsBrowse.Location = new Point(329, prevBtnY);
            btnGapsBrowse.Click += btnGapsBrowse_Click;
            this.Controls.Add(btnGapsBrowse);

            //Gaps Textbox
            txtGapsPath = new TextBox();
            txtGapsPath.Width = 244;
            txtGapsPath.Height = 20;
            txtGapsPath.Name = "txtGapsPath";
            txtGapsPath.Leave += txtGapsPath_Leave;
            txtGapsPath.Location = new Point(btnGapsBrowse.Location.X - 250, btnGapsBrowse.Location.Y + 2);
            txtGapsPath.Text = Properties.Settings.Default.PathGAPS;
            this.Controls.Add(txtGapsPath);

            //Gaps Label
            lblGapsPath = new Label();
            lblGapsPath.Text = "GAPS Path";
            lblGapsPath.Name = "lblGapsPath";
            this.Controls.Add(lblGapsPath);
            lblGapsPath.AutoSize = true;
            lblGapsPath.Location = new Point(txtGapsPath.Location.X - lblGapsPath.Width - 7, txtGapsPath.Location.Y + 3);
            #endregion

            #region 117
            //117 Button
            btn117Browse = new Button();
            btn117Browse.Text = "Browse";
            btn117Browse.Name = "btn117Browse";
            btn117Browse.Location = new Point(329, prevBtnY += 26);
            btn117Browse.Click += btn117Browse_Click;
            this.Controls.Add(btn117Browse);

            //117 Textbox
            txt117Path = new TextBox();
            txt117Path.Width = 244;
            txt117Path.Height = 20;
            txt117Path.Name = "txt117Path";
            txt117Path.Location = new Point(btn117Browse.Location.X - 250, btn117Browse.Location.Y + 2);
            txt117Path.Text = Properties.Settings.Default.Path117;
            this.Controls.Add(txt117Path);

            //117 Label
            lbl117Path = new Label();
            lbl117Path.Text = "117 Path";
            lbl117Path.Name = "lbl117Path";
            this.Controls.Add(lbl117Path);
            lbl117Path.AutoSize = true;
            lbl117Path.Location = new Point(txt117Path.Location.X - lbl117Path.Width - 8, txt117Path.Location.Y + 3);
            #endregion

            #region 473
            //473 Button
            btn473Browse = new Button();
            btn473Browse.Text = "Browse";
            btn473Browse.Name = "btn473Browse";
            btn473Browse.Location = new Point(329, prevBtnY += 26);
            btn473Browse.Click += btn473Browse_Click;
            this.Controls.Add(btn473Browse);

            //473 Textbox
            txt473Path = new TextBox();
            txt473Path.Width = 244;
            txt473Path.Height = 20;
            txt473Path.Name = "txt473Path";
            txt473Path.Location = new Point(btn473Browse.Location.X - 250, btn473Browse.Location.Y + 2);
            txt473Path.Text = Properties.Settings.Default.Path473;
            this.Controls.Add(txt473Path);

            //473 Label
            lbl473Path = new Label();
            lbl473Path.Text = "473 Path";
            lbl473Path.Name = "lbl473Path";
            this.Controls.Add(lbl473Path);
            lbl473Path.AutoSize = true;
            lbl473Path.Location = new Point(txt473Path.Location.X - lbl473Path.Width - 8, txt473Path.Location.Y + 3);
            #endregion

            #region 325
            //325 Button
            btn325Browse = new Button();
            btn325Browse.Text = "Browse";
            btn325Browse.Name = "btn325Browse";
            btn325Browse.Location = new Point(329, prevBtnY += 26);
            btn325Browse.Click += btn325Browse_Click;
            this.Controls.Add(btn325Browse);

            //325 Textbox
            txt325Path = new TextBox();
            txt325Path.Width = 244;
            txt325Path.Height = 20;
            txt325Path.Name = "txt325Path";
            txt325Path.Location = new Point(btn325Browse.Location.X - 250, btn325Browse.Location.Y + 2);
            txt325Path.Text = Properties.Settings.Default.Path325;
            this.Controls.Add(txt325Path);

            //325 Label
            lbl325Path = new Label();
            lbl325Path.Text = "325 Path";
            lbl325Path.Name = "lbl325Path";
            this.Controls.Add(lbl325Path);
            lbl325Path.AutoSize = true;
            lbl325Path.Location = new Point(txt325Path.Location.X - lbl325Path.Width - 8, txt325Path.Location.Y + 3);
            #endregion

            #region AP1000
            //AP1000 Button
            btnAP1000Browse = new Button();
            btnAP1000Browse.Text = "Browse";
            btnAP1000Browse.Name = "btnAP1000Browse";
            btnAP1000Browse.Location = new Point(329, prevBtnY += 26);
            btnAP1000Browse.Click += btnAP1000Browse_Click;
            this.Controls.Add(btnAP1000Browse);

            //AP1000 Textbox
            txtAP1000Path = new TextBox();
            txtAP1000Path.Width = 244;
            txtAP1000Path.Height = 20;
            txtAP1000Path.Name = "txtAP1000Path";
            txtAP1000Path.Location = new Point(btnAP1000Browse.Location.X - 250, btnAP1000Browse.Location.Y + 2);
            txtAP1000Path.Text = Properties.Settings.Default.PathAP1000;
            this.Controls.Add(txtAP1000Path);

            //AP1000 Label
            lblAP1000Path = new Label();
            lblAP1000Path.Text = "AP1000 Path";
            lblAP1000Path.Name = "lblAP1000Path";
            this.Controls.Add(lblAP1000Path);
            lblAP1000Path.AutoSize = true;
            lblAP1000Path.Location = new Point(txtAP1000Path.Location.X - lblAP1000Path.Width - 8, txtAP1000Path.Location.Y + 3);
            #endregion

            #region FormHeight
            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    txtCounter += 1;
                }
            }

            this.Height = (txtCounter * 20) + (6 * (txtCounter - 1)) + 87;
            #endregion

            #region SaveBtn
            btnSave = new Button();
            btnSave.Text = "Save";
            btnSave.Name = "btnSave";
            btnSave.Location = new Point(248, this.Height - 59);
            btnSave.Click += btnSave_Click;
            this.Controls.Add(btnSave);
            #endregion

            #region CancelBtn
            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Name = "btnCancel";
            btnCancel.Location = new Point(329, this.Height - 59);
            btnCancel.Click += btnCancel_Click;
            this.Controls.Add(btnCancel);
            #endregion
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
