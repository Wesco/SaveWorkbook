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
        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.SaveReport();
        }

        private void btnVMI_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.SaveVMI();
        }

        private void btnSaveOAR_Click(object sender, RibbonControlEventArgs e)
        {
            App.thisAddin.SaveOAR();
        }

        private void btnConfigure_Click(object sender, RibbonControlEventArgs e)
        {
            frmSettings frm = new frmSettings();
            frm.Show();
        }
    }
}
