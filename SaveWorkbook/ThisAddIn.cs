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
            string[] reportType = new string[2];
            string branch = "";
            int ISN = 0;

            //Filter the report type string to check if it is a back order report
            if (!String.IsNullOrEmpty(ActiveSheet.Range["A1"].Value))
            {
                branch = (ActiveSheet.Range["A3"].Value).ToString();

                reportType[0] = (ActiveSheet.Range["A1"].Value).ToString();
                reportType[1] = (ActiveSheet.Range["A1"].Value).ToString();


                reportType[0] = reportType[0].Replace(" ", String.Empty);
                reportType[0] = reportType[0].Substring(reportType[0].Length - 10);

                reportType[1] = reportType[1].Replace(" ", String.Empty);
                reportType[1] = reportType[1].Substring(reportType[1].Length - 19);
            }

            //If it is a back order report
            if (reportType[0] == "BACKORDERS")
            {
                //Try to find the inside sales number
                int.TryParse((ActiveSheet.Cells[3, FindColumnHeader(2, "IN")].Value).ToString(), out ISN);

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
            else if (reportType[1] == "DIRECTSHIPPEDORDERS")
            {
                //Try to find the inside sales number
                int.TryParse((ActiveSheet.Cells[3, FindColumnHeader(2, "IN")].Value).ToString(), out ISN);

                if (ISN != 0)
                {
                    path = @"\\br3615gaps\gaps\3615 117 Report\ByInsideSalesNumber\" + ISN + @"\";
                    fileName = branch + " " + String.Format("{0:M-dd-yy}", dt) + " DSORDERS" + ".xlsx";

                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    try
                    {
                        App.oApp.ActiveWorkbook.SaveAs(path + fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    catch (Exception e)
                    {
                        //If error is not due to user canceled save display the error message
                        if (e.Message.ToLower() != "exception from hresult: 0x800a03ec")
                            System.Windows.Forms.MessageBox.Show(e.Message.ToString());
                        //System.Windows.Forms.MessageBox.Show("File could not be saved!");
                    }
                }

            }
        }

        /// <summary>
        /// Finds the specified column and returns the column number.
        /// If no column is found then 0 is returned.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        private int FindColumnHeader(int row, string text)
        {
            for (int col = 1; col < ActiveSheet.UsedRange.Columns.Count; col++)
            {
                if ((ActiveSheet.Cells[row, col].Value).ToString() == text)
                {
                    return col;
                }
            }

            return 0;
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
