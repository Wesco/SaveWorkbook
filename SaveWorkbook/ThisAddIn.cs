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
    public class App
    {
        public static ThisAddIn thisAddin { get; set; }
    }

    static class Extensions
    {
        public static string Right(this string value, int length)
        {
            return value.Substring(value.Length - length);
        }

        public static string Left(this string value, int length)
        {
            return value.Substring(0, length);
        }

        public static string Find(this string value, string text)
        {
            int index = 0;

            index = value.IndexOf(text, 0);

            if (index > 0)
                return value.Substring(index, text.Length);
            else
                return String.Empty;
        }
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
        private Excel.Workbook ActiveWorkbook { get; set; }
        private Excel.Sheets Sheets { get; set; }
        private Excel.Workbooks Workbooks { get; set; }
        private Excel.Workbook ThisWorkbook { get; set; }

        public void SaveReport()
        {
            string reptype;

            if (!String.IsNullOrEmpty(ActiveSheet.Range["A1"].Value))
            {
                reptype = ((ActiveSheet.Range["A1"].Value).ToString().Replace(" ", String.Empty)).Substring(0, 3);

                switch (reptype)
                {
                    case "117":
                        string type = (ActiveSheet.Range["A1"].Value).ToString().Replace(" ", String.Empty);
                        if (type.Find("BYINSIDESALESPERSON") == "BYINSIDESALESPERSON")
                            SaveISN117();
                        break;

                    case "473":
                        Save473();
                        break;

                    case "Branch_id":
                        SaveGAPS();
                        break;

                    default:
                        System.Windows.Forms.MessageBox.Show("This report is not handled by this add-in.");
                        break;
                }
            }
        }

        public void SaveISN117()
        {
            DateTime dt = DateTime.Now;

            string path;
            string fileName;
            string[] reportType = new string[3];
            string branch = "";
            int ISN = 0;

            //Filter the report type string to check if it is a back order report
            if (!String.IsNullOrEmpty(ActiveSheet.Range["A1"].Value))
            {
                branch = (ActiveSheet.Range["A3"].Value).ToString();
                reportType[0] = (ActiveSheet.Range["A1"].Value).ToString();
                reportType[1] = (ActiveSheet.Range["A1"].Value).ToString();
                reportType[2] = (ActiveSheet.Range["A1"].Value).ToString();

                reportType[0] = reportType[0].Replace(" ", String.Empty);
                reportType[0] = reportType[0].Substring(reportType[0].Length - 10);

                reportType[1] = reportType[1].Replace(" ", String.Empty);
                reportType[1] = reportType[1].Substring(reportType[1].Length - 19);

                reportType[2] = reportType[0].Replace(" ", String.Empty);
                reportType[2] = reportType[0].Substring(reportType[0].Length - 9);
            }

            for (int i = 0; i < reportType.Count(); i++)
            {
                switch (reportType[i])
                {
                    case "BACKORDERS":
                        //Try to find the inside sales number
                        int.TryParse((ActiveSheet.Cells[3, FindColumnHeader(2, "IN")].Value).ToString(), out ISN);

                        if (ISN != 0)
                        {
                            fileName = branch + " " + String.Format("{0:M-dd-yy}", dt) + " BACKORDERS" + ".xlsx";
                            path = Properties.Settings.Default.Path117 + @"ByInsideSalesNumber\" + ISN + @"\";
                            if (!Directory.Exists(path))
                                Directory.CreateDirectory(path);

                            SaveActiveBook(path, fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                        }
                        break;

                    case "DIRECTSHIPPEDORDERS":
                        //Try to find the inside sales number
                        int.TryParse((ActiveSheet.Cells[3, FindColumnHeader(2, "IN")].Value).ToString(), out ISN);

                        if (ISN != 0)
                        {
                            path = Properties.Settings.Default.Path117 + @"ByInsideSalesNumber\" + ISN + @"\";
                            fileName = branch + " " + String.Format("{0:M-dd-yy}", dt) + " DSORDERS" + ".xlsx";

                            if (!Directory.Exists(path))
                                Directory.CreateDirectory(path);

                            SaveActiveBook(path, fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                        }
                        break;

                    case "ALLORDERS":
                        //Try to find the inside sales number
                        int.TryParse((ActiveSheet.Cells[3, FindColumnHeader(2, "IN")].Value).ToString(), out ISN);

                        if (ISN != 0)
                        {
                            path = Properties.Settings.Default.Path117 + @"ByInsideSalesNumber\" + ISN + @"\";
                            fileName = branch + " " + String.Format("{0:M-dd-yy}", dt) + " ALLORDERS" + ".xlsx";

                            if (!Directory.Exists(path))
                                Directory.CreateDirectory(path);

                            SaveActiveBook(path, fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                        }
                        break;
                }
            }
        }

        public void Save473()
        {
            DateTime dt = DateTime.Now;
            string fileName = "473 " + string.Format("{0:M-dd-yy}", dt) + ".xlsx";
            string reportType = (ActiveSheet.Range["A1"].Value).ToString();
            string branch = (ActiveSheet.Range["A3"].Value).ToString();
            string path = @"\\br3615gaps\gaps\" + branch + @" 473 Download\";

            //Verify that report is a 473 Open Order Report
            if (reportType.Length > 3)
            {
                reportType = reportType.Substring(0, 3);

                //If it is a 473 then try to save
                if (reportType == "473")
                {
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);

                    try
                    {
                        ActiveWorkbook.SaveAs(path + fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    catch (Exception e)
                    {
                        //If error is not due to user canceled save display the error message
                        if (e.Message.ToLower() != "exception from hresult: 0x800a03ec")
                            System.Windows.Forms.MessageBox.Show(e.Message.ToString());
                    }
                }
            }
        }

        public void SaveVMI()
        {
            foreach (Excel.Worksheet s in ActiveWorkbook.Worksheets)
            {
                DateTime dt = DateTime.Now.AddMonths(-1);
                Office.MsoFileDialogType dlgType = Office.MsoFileDialogType.msoFileDialogSaveAs;

                if (s.Name != "Drop In" &&
                    s.Name != "PivotTable" &&
                    s.Name != "Info" &&
                    s.Name != "Macro" &&
                    s.Name != "VMI eStock" &&
                    s.Name != "Master")
                {
                    s.Copy();
                    if (s.Range["C5"].Text != "")
                    {
                        s.Columns[ActiveSheet.UsedRange.Columns.Count].Delete();
                    }
                    Application.FileDialog[dlgType].InitialFileName = s.Name + "_" + String.Format("{0:MMM_yyyy}", dt);
                    Application.FileDialog[dlgType].Show();
                    if (Application.FileDialog[dlgType].SelectedItems.Count > 0)
                    {
                        ActiveWorkbook.SaveAs(Application.FileDialog[dlgType].SelectedItems.Item(1), Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    Application.DisplayAlerts = false;
                    ActiveWorkbook.Close();
                    Application.DisplayAlerts = true;
                }
            }

        }

        public void SaveGAPS()
        {
            DateTime dt = DateTime.Now;
            string Branch = (ActiveSheet.Range["A2"].Value).ToString();
            string sPath = @"\\BR3615GAPS\Gaps\" + Branch + @" Gaps Download\" + String.Format("{0:yyyy}", dt) + @"\";
            string sFile = Branch + " " + String.Format("{0:M-dd-yy}", dt) + ".xlsx";

            try
            {
                //Try to verify that the file being saved is actuall GAPS
                if ((ActiveSheet.Range["A1"].Value).ToString() == "Branch_id" && (ActiveSheet.Range["CU1"].Value).ToString() == "Wdc_rt_qty")
                    ActiveWorkbook.SaveAs(sPath + sFile, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception e)
            {
                //If error is not due to user canceled save display the error message
                if (e.Message.ToLower() != "exception from hresult: 0x800a03ec")
                    System.Windows.Forms.MessageBox.Show(e.Message.ToString());
            }
        }

        /// <summary>
        /// Finds the specified column and returns the column number.
        /// If no column is found then 0 is returned.
        /// </summary>
        /// <param name="row">The row containg column headers</param>
        /// <param name="text">The column header to search for</param>
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

        private void SaveActiveBook(string Path, string FileName, Excel.XlFileFormat FileFormat)
        {
            try
            {
                ActiveWorkbook.SaveAs(Path + FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception e)
            {
                //If error is not due to user canceled save display the error message
                if (e.Message.ToLower() != "exception from hresult: 0x800a03ec")
                    System.Windows.Forms.MessageBox.Show(e.Message.ToString());
            }
        }

        #region Event_Handlers
        void Application_SheetActivate(object Sh)
        {
            ActiveSheet = Application.ActiveSheet;
        }

        void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            ActiveWorkbook = Application.ActiveWorkbook;
            ActiveSheet = Application.ActiveSheet;
            Sheets = Application.ActiveWorkbook.Sheets;
        }

        void Application_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            Sheets = Application.ActiveWorkbook.Sheets;
            Workbooks = Application.Workbooks;
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            ActiveWorkbook = Application.ActiveWorkbook;
            ActiveSheet = Application.ActiveSheet;
            Sheets = Application.ActiveWorkbook.Sheets;
            Workbooks = Application.Workbooks;
        }
        #endregion

        #region AddIn_Events

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            App.thisAddin = this;
            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.SheetActivate += Application_SheetActivate;
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            ActiveWorkbook = Application.ActiveWorkbook;
            ActiveSheet = Application.ActiveSheet;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.SheetActivate -= Application_SheetActivate;
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
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
}
