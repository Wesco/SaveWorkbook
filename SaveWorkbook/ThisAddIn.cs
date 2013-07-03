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

            if (index >= 0)
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

            if (ActiveSheet.Range["A1"].Value != null)
            {
                reptype = ((ActiveSheet.Range["A1"].Value).ToString().Replace(" ", String.Empty)).Substring(0, 3);

                switch (reptype)
                {
                    case "117":
                        string type = (ActiveSheet.Range["A1"].Value).ToString().Replace(" ", String.Empty);
                        if (type.Find("BYINSIDESALESPERSON") == "BYINSIDESALESPERSON")
                            SaveISN117();
                        if (type.Find("DETAILREPORTBYCUSTOMER") == "DETAILREPORTBYCUSTOMER")
                            SaveCust117();
                        break;

                    case "473":
                        Save473();
                        break;

                    case "Bra":
                        if (IsGaps())
                            SaveGAPS();
                        break;

                    case "SIM":
                        Save325();
                        break;

                    default:
                        System.Windows.Forms.MessageBox.Show("This report is not handled by this add-in.");
                        break;
                }
            }
        }

        public void SaveCust117()
        {
            string type = (ActiveSheet.Range["A1"].Value).ToString();
            string branch = (ActiveSheet.Range["A3"].Value).ToString();
            string custNum = (ActiveSheet.Range["C3"].Value).ToString();
            string fileName = branch + " " + Today() + " INQUIRY";
            string savePath = Properties.Settings.Default.Path117 + branch + " 117 Report\\" + "ByCustomerNumber\\" + custNum + "\\";

            if (type.Find("INQUIRIES") == "INQUIRIES")
            {
                SaveActiveBook(savePath, fileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
        }

        public void SaveISN117()
        {
            string path;
            string fileName;
            string[] reportType = new string[3];
            string branch = "";
            int ISN = 0;

            //Filter the report type string to check if it is a back order report
            if (ActiveSheet.Range["A1"].Value != null)
            {
                if (ActiveSheet.Range["A3"].Value != null)
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
                            fileName = branch + " " + Today() + " BACKORDERS" + ".xlsx";
                            path = Properties.Settings.Default.Path117 + branch + @" 117 Report\" + @"ByInsideSalesNumber\" + ISN + @"\";
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
                            path = Properties.Settings.Default.Path117 + branch + @" 117 Report\" + @"ByInsideSalesNumber\" + ISN + @"\";
                            fileName = branch + " " + Today() + " DSORDERS" + ".xlsx";

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
                            path = Properties.Settings.Default.Path117 + branch + @" 117 Report\" + @"ByInsideSalesNumber\" + ISN + @"\";
                            fileName = branch + " " + Today() + " ALLORDERS" + ".xlsx";

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
            string fileName = "473 " + Today() + ".xlsx";
            string reportType = (ActiveSheet.Range["A1"].Value).ToString();
            string branch = (ActiveSheet.Range["A3"].Value).ToString();
            string path = Properties.Settings.Default.Path473 + branch + @" 473 Download\";

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
            string Branch = (ActiveSheet.Range["A2"].Value).ToString();
            string sPath = Properties.Settings.Default.PathGAPS + Branch + @" Gaps Download\" + Today() + @"\";
            string sFile = Branch + " " + Today() + ".xlsx";

            if (!Directory.Exists(sPath))
                Directory.CreateDirectory(sPath);

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

        public void Save325()
        {
            string path = Properties.Settings.Default.Path325;
            string fileName = "325 " + Today() + ".xlsx";
            string reportType = ActiveSheet.Range["A1"].Value;

            if (reportType.Left(7) == "SIMLIST" && reportType.Replace(" ", String.Empty).Right(17) == "INVENTORYDOWNLOAD")
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
            if (!Directory.Exists(Path))
            {
                try
                {
                    Directory.CreateDirectory(Path);
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.Message.ToString());
                }
            }

            try
            {
                if (Directory.Exists(Path))
                {
                    Application.DisplayAlerts = false;
                    ActiveWorkbook.SaveAs(Path + FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    Application.DisplayAlerts = true;
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                //If error is not due to user canceled save display the error 
                if (e.Message.ToLower() != "exception from hresult: 0x800a03ec")
                    System.Windows.Forms.MessageBox.Show(e.Message.ToString());
            }
        }

        private bool IsGaps()
        {
            #region List of GAPS Headers
            string[] GapsHeaders = new string[99]
            {
                "Branch_id",
                "Sim_mfr_no",
                "Sim_item_no",
                "buyerCode",
                "Sim_description",
                "Qty_on_hand",
                "Qty_on_reserve",
                "Qty_on_backorder",
                "Qty_on_order",
                "qty_on_consignment",
                "Mtd_sales_qty",
                "WESCOM_SLS1",
                "WESCOM_SLS2",
                "WESCOM_SLS3",
                "WESCOM_SLS4",
                "WESCOM_SLS5",
                "WESCOM_SLS6",
                "WESCOM_SLS7",
                "WESCOM_SLS8",
                "WESCOM_SLS9",
                "WESCOM_SLS10",
                "WESCOM_SLS11",
                "WESCOM_SLS12",
                "Order_review_point",
                "Fixed_review_point",
                "Basic_stock",
                "Fixed_basic_stock",
                "master_stock",
                "WESCOM_QTY_BREAK_2",
                "WESCOM_QTY_BREAK_3",
                "Last_cost",
                "Unit_Last_cost",
                "average_cost",
                "Unit_average_cost",
                "Unit_of_measure_id",
                "Wdc_qty_on_hand",
                "Product_code",
                "Supplier_no",
                "WESCOM_Buyer_Code",
                "Inventory_as_of",
                "Date_Last_Changed",
                "Date_Net_Stock_Decrease",
                "Item_Note",
                "DSS_data_as_of",
                "country_code",
                "Velocity_code",
                "Class_code",
                "WESCOM_Leadtime",
                "start_date",
                "end_date",
                "lead_time",
                "value_code",
                "last_12_sls_qty",
                "max_qty",
                "one_ms_qty",
                "rec_target_qty",
                "except_status",
                "sls_qty1",
                "sls_qty2",
                "sls_qty3",
                "sls_qty4",
                "sls_qty5",
                "sls_qty6",
                "sls_qty7",
                "sls_qty8",
                "sls_qty9",
                "sls_qty10",
                "sls_qty11",
                "sls_qty12",
                "sls_qty13",
                "sls_qty14",
                "sls_qty15",
                "sls_qty16",
                "sls_qty17",
                "sls_qty18",
                "sls_qty19",
                "sls_qty20",
                "sls_qty21",
                "sls_qty22",
                "sls_qty23",
                "sls_qty24",
                "Forecast1",
                "Forecast2",
                "Forecast3",
                "Replacement_qty_break2",
                "Replacement_qty_break3",
                "mfr_name",
                "Min_Purchase",
                "Min_Freight",
                "contact",
                "phone",
                "fax",
                "rec_days",
                "def_rec_days",
                "Date_First_Receipt",
                "Date_last_issued",
                "Date_Created",
                "Days_Supply",
                "Wdc_rt_qty",
            };
            #endregion
            int TotalCols = ActiveSheet.UsedRange.Columns.Count;
            Excel.Range ReportHeaders = ActiveSheet.Range[ActiveSheet.Cells[1, 1], ActiveSheet.Cells[1, TotalCols]];

            if (ReportHeaders.Columns.Count == GapsHeaders.Length)
                for (int i = 0; i < GapsHeaders.Length; i++)
                {
                    if (ReportHeaders.Cells[1, i + 1].Value != GapsHeaders[i])
                        return false;
                }
            else
                return false;

            return true;
        }

        /// <summary>
        /// Gets todays date and converts it into ISO 8601 compliant string.
        /// </summary>
        /// <returns>Todays date in yyyy-MM-dd format</returns>
        private string Today()
        {
            DateTime dt = DateTime.Now;
            string date = String.Format("{0:yyyy-MM-dd}", dt);
            return date;
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
