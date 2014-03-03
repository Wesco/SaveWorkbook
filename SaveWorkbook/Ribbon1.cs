using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace SaveWorkbook
{
    public partial class rbnSaveReport
    {
        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            oApp.thisAddin.SaveReport();
        }

        private void btnVMI_Click(object sender, RibbonControlEventArgs e)
        {
            oApp.thisAddin.SaveVMI();
        }

        private void btnConfigure_Click(object sender, RibbonControlEventArgs e)
        {
            frmSettings frm = new frmSettings();

            if (Application.OpenForms[frm.Name] == null)
                frm.Show();
            else
                Application.OpenForms[frm.Name].Focus();
        }
    }
}
