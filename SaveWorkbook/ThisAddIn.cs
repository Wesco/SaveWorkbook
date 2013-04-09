using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;

namespace SaveWorkbook
{
    public class aApp
    {
        public static ThisAddIn thisAddin { get; set; }
    }

    public partial class ThisAddIn
    {
        private Excel.Worksheet sheet;
        private Excel.Worksheet ActiveSheet
        {
            get
            {
                if (sheet != null)
                {
                    return sheet;
                }
                else
                {
                    return Application.ActiveSheet;
                }
            }

            set
            {
                sheet = value;
            }
        }

        //Save 117 by inside sales number
        public void SaveISN117()
        {
            DateTime dt = DateTime.Now;

            string path;
            string fileName;
            string reportType = "";
            string branch = "";
            int ISN = 0;

            //Filter the report type string to check if it is a back order report
            if (!String.IsNullOrEmpty(ActiveSheet.Range["A1"].Value))
            {
                branch = (ActiveSheet.Range["A3"].Value).ToString();

                reportType = (ActiveSheet.Range["A1"].Value).ToString();
                reportType = reportType.Replace(" ", String.Empty);
                reportType = reportType.Substring(reportType.Length - 10);
            }

            //If it is a back order report
            if (reportType == "BACKORDERS")
            {
                //Try to find the inside sales number
                for (int i = 1; i < ActiveSheet.UsedRange.Columns.Count; i++)
                {
                    if ((string)ActiveSheet.Cells[2, i].Value == "IN")
                    {
                        ISN = (int)ActiveSheet.Cells[3, i].Value;
                        break;
                    }
                }

                if (ISN != 0)
                {
                    path = @"\\br3615gaps\gaps\3615 117 Report\ByInsideSalesNumber\" + ISN + @"\";
                    fileName = branch + " " + String.Format("{0:M-dd-yy}", dt) + " BACKORDERS" + ".xlsx";

                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    try
                    {
                        App.oApp.ActiveWorkbook.SaveAs(path + fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    catch (Exception)
                    {
                        System.Windows.Forms.MessageBox.Show("File could not be saved!");
                    }
                }
            }

            System.Windows.Forms.MessageBox.Show(reportType);
        }


        #region AddIn_Events
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            App.oApp = this.Application;
            aApp.thisAddin = this;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    public class App
    {
        public static Excel.Application oApp { get; set; }
    }
}
