using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace SaveWorkbook
{
    public partial class rbnSaveReport
    {
        private void btnGaps_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.SaveGAPS();
        }

        private void btnISN117_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.SaveISN117();
        }

        private void btn473_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.Save473();
        }

        private void btnVMI_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.SaveVMI();
        }

        private void btnConfigure_Click(object sender, RibbonControlEventArgs e)
        {
            frmSettings frm = new frmSettings();
            frm.Show();
        }
    }
}
