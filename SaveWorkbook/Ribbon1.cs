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
            SaveWkbk((App.oApp.ActiveWorkbook.ActiveSheet.Range("A2").Value).ToString());
        }

        private void SaveWkbk(String Branch)
        {
            DateTime dt = DateTime.Now;
            string sPath = @"\\BR3615GAPS\Gaps\" + Branch + @" Gaps Download\" + String.Format("{0:yyyy}", dt) + @"\";
            string sFile = Branch + " " + String.Format("{0:M-dd-yy}", dt) + ".xlsx";

            try
            {
                App.oApp.ActiveWorkbook.SaveAs(sPath + sFile, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("File not saved!");
            }
        }
    }
}
