using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using MessageBox = System.Windows.Forms.MessageBox;
using System.Globalization;

namespace SaveWorkbook
{
    public class App
    {
        public static ThisAddIn thisAddin { get; set; }
    }

    public partial class ThisAddIn
    {
        private Excel.Worksheet activeSheet;
        private Excel.Workbook activeBook;

        private Excel.Worksheet ActiveSheet
        {
            get
            {
                if (activeSheet != null)
                {
                    return activeSheet;
                }
                else
                {
                    return Application.ActiveSheet;
                }
            }

            set
            {
                activeSheet = value;
            }
        }
        private Excel.Workbook ActiveWorkbook
        {
            get
            {
                if (activeBook != null)
                {
                    return activeBook;
                }
                else
                {
                    return Application.ActiveWorkbook;
                }
            }

            set
            {
                activeBook = value;
            }
        }
        private Excel.Sheets Sheets { get; set; }
        private Excel.Workbooks Workbooks { get; set; }
        private Excel.Workbook ThisWorkbook { get; set; }

        private const string RepNotHandled = "This report is not handled by this add-in.";

        public void SaveReport()
        {
            string reptype;

            if (ActiveSheet.Range["A1"].Value != null)
            {
                int condition = (ActiveSheet.Range["A1"].Value).ToString().Length;

                if (condition > 2)
                    reptype = ((ActiveSheet.Range["A1"].Value).ToString().Replace(" ", String.Empty)).Substring(0, 3);
                else
                    reptype = String.Empty;

                switch (reptype)
                {
                    case "117":
                        if (Is117())
                            Save117();
                        else
                            MessageBox.Show(RepNotHandled, "RepType 117 - Error");
                        break;

                    case "473":
                        if (Is473())
                            Save473();
                        break;

                    case "Bra":
                        if (IsGaps())
                            SaveGAPS();
                        break;

                    case "SIM":
                        Save325();
                        break;

                    case "Sup":
                        if (IsIROOR())
                            SaveIROpenOrders();
                        break;

                    default:
                        MessageBox.Show(RepNotHandled);
                        break;
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The report type could not be determined.");
            }
        }

        #region 117
        public void Save117()
        {
            TextInfo txtInfo = new CultureInfo("en-US", false).TextInfo;
            string Branch = GetString(ActiveSheet.Range["A3"]);
            string Identifier = GetString(ActiveSheet.Range["A1"]).SingleSpace().Trim();
            string FileName = Branch + " " + Today();
            string SavePath = Properties.Settings.Default.Path117 + Branch + " 117 Report\\";
            string Criteria = String.Empty;
            string Sequence = String.Empty;
            string DetailSummary = String.Empty;
            int ByIndex = 0;
            int ForIndex = 0;

            // Verify the branch number was found
            if (Branch == String.Empty)
            {
                MessageBox.Show("Unable to find branch number.", "Save117 Error - Branch");
                return;
            }

            // Check if report is a detailed or summary report
            if (Identifier.Contains("SUMMARY REPORT"))
                DetailSummary = "SUMMARY";
            else if (Identifier.Contains("DETAIL REPORT"))
                DetailSummary = "DETAIL";
            else
            {
                MessageBox.Show(RepNotHandled);
                return;
            }

            // Add Detail/Summary to the file path
            SavePath += DetailSummary + "\\";

            // Get report sequence
            ByIndex = Identifier.IndexOf("BY");
            ForIndex = Identifier.IndexOf("FOR");

            try
            {
                Sequence = txtInfo.ToTitleCase(Identifier.Substring(ByIndex, ForIndex - ByIndex).ToLower()).Replace(" ", "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Report Sequence - Error");
                return;
            }

            // Get report criteria
            try
            {
                Criteria = Identifier.Right(Identifier.Length - ForIndex - 4);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Report Criteria - Error");
                return;
            }

            // Sequence = Report Sequence
            // Used to determine the save folder
            #region Sequence
            switch (Sequence)
            {
                #region ByOrder
                case "ByOrder":
                    SavePath += Sequence + "\\";
                    break;
                #endregion

                #region ByCustomer
                case "ByCustomer":
                    //TODO:
                    //Check to see if a range of customers was chosen
                    //If more than one DPC is listed return as not handled
                    int ColNum = FindColumnHeader(2, "CUSTOMER");
                    string Customer = GetString(ActiveSheet.Cells[3, ColNum]);
                    if (Customer != String.Empty)
                        SavePath += Sequence + "\\" + Customer + "\\";
                    else
                    {
                        MessageBox.Show("Unable to find customer DPC.", "Sequence ByCustomer - Error");
                        return;
                    }
                    break;
                #endregion

                // By Order Date not handled
                // By SIM Number not handled
                // By Gross Margin not handled
                // By Dollar Amount not handled

                #region ByInsideSalesperson
                case "ByInsideSalesperson":
                    //TODO:
                    //Add a check to make sure only one sales number was chosen
                    //If multiple were chosen return as not handled
                    string ISN = GetString(ActiveSheet.Cells[3, FindColumnHeader(2, "IN")]);
                    if (ISN != String.Empty)
                        SavePath += Sequence + "\\" + ISN + "\\";
                    else
                    {
                        MessageBox.Show("Unable to find inside sales number.", "Sequence ByISN - Error");
                        return;
                    }
                    break;
                #endregion

                #region ByOutsideSalesperson
                case "ByOutsideSalesperson":
                    //TODO:
                    //Add a check to make sure only one sales number was chosen
                    //If multiple were chosen return as not handled
                    string OSN = GetString(ActiveSheet.Cells[3, FindColumnHeader(2, "OUT")]);
                    if (OSN != String.Empty)
                        SavePath += Sequence + "\\" + OSN + "\\";
                    else
                    {
                        MessageBox.Show("Unable to find outside sales number.", "Sequence ByOSN - Error");
                        return;
                    }
                    break;
                #endregion

                #region Default
                default:
                    MessageBox.Show(RepNotHandled, "Sequence Default - Error");
                    return;
                #endregion
            }
            #endregion

            // Criteria = Report Selection Critieria
            // Used to determine the file name
            #region Criteria
            switch (Criteria)
            {
                case "ALL ORDERS":
                    FileName += " ALLORDERS";
                    break;

                case "BACK ORDERS":
                    FileName += " BACKORDERS";
                    break;

                // Contract Orders not handled

                case "DIRECT SHIPPED ORDERS":
                    FileName += " DSORDERS";
                    break;

                case "INQUIRIES":
                    FileName += " INQUIRIES";
                    break;

                case "CREDIT MEMOS":
                    FileName += "CREDITMEMOS";
                    break;

                // New Orders (Entered Not Picked) not handled

                case "OPEN PICK TICKETS":
                    FileName += " OPENTICKETS";
                    break;

                case "ORDERS SHIPPED BUT NOT INVOICED":
                    FileName += " SHIPPEDNOTINVOICED";
                    break;

                case "UNRELEASED ORDERS":
                    FileName += " UNRELEASED";
                    break;

                // Blanket Orders Held Pendign Review not handled

                case "SPECIAL ORDERS":
                    FileName += " SPECIALORDERS";
                    break;

                case "ASSEMBLE AND HOLD ORDERS":
                    FileName += " ASSEMBLEHOLD";
                    break;

                default:
                    MessageBox.Show(RepNotHandled, "Criteria Default - Error");
                    return;
            }
            #endregion

            SaveActiveBook(SavePath, FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
        }

        private bool Is117()
        {
            string[] ColHeaders = new string[0];
            string Identifier = String.Empty;
            int TotalCols = ActiveSheet.UsedRange.Columns.Count;
            Excel.Range ReportHeaders = ActiveSheet.Range[ActiveSheet.Cells[1, 1], ActiveSheet.Cells[1, TotalCols]];

            //Get the report identifier string in A1
            if (ActiveSheet.Range["A1"].Value != null)
            {
                Identifier = (ActiveSheet.Range["A1"].Value).ToString();
                Identifier = Identifier.Replace(" ", String.Empty);
                Identifier = Identifier.Replace("\t", String.Empty);
            }
            else
                return false;

            //Check the identifier string for the report type
            if (Identifier.Contains("SUMMARYREPORT") && Identifier.Contains("117"))
                #region Summary Column Headers
                ColHeaders = new string[21]
                {
                    "WAREHOUSE",
                    "ERROR",
                    "CUSTOMER",
                    "SHIP TO",
                    "CUSTOMER NAME",
                    "CUSTOMER REFERENCE NO",
                    "CUSTOMER PART NUMBER",
                    "ORDER DATE",
                    "ORDER NO",
                    "CYCLE",
                    "TYPE",
                    "STATUS",
                    "SHIP COMPLETE",
                    "ORDER AMOUNT",
                    "IN",
                    "OUT",
                    "STATUS DESCRIPTION",
                    "SUSPENSION TYPE",
                    "EXT COST",
                    "EXT MARGIN $",
                    "QUOTED TO"
                };
                #endregion
            else if (Identifier.Contains("DETAILREPORT") && Identifier.Contains("117"))
                #region DetailColHeaders
                ColHeaders = new string[62] 
                {
                    "WAREHOUSE",
                    "ERROR",
                    "CUSTOMER",
                    "ORDER NO",
                    "REMOTE ORDER",
                    "CYCLE",
                    "STATUS",
                    "ORDER DATE",
                    "TAX",
                    "TAX ACCOUNT",
                    "REQUIRED DATE (HR)",
                    "CUSTOMER REFERENCE NO",
                    "CUST PO LINE #",
                    "CUSTOMER PART NUMBER",
                    "SHIP TO",
                    "IN",
                    "OUT",
                    "LINE NO",
                    "KIT",
                    "TYPE",
                    "ITEM NUMBER",
                    "CATALOG NUMBER",
                    "ITEM DESCRIPTION",
                    "SUOM",
                    "ORDER QTY",
                    "AVAILABLE QTY",
                    "QTY TO SHIP",
                    "BO QTY",
                    "QTY SHIPPED",
                    "GROSS MARGIN",
                    "LPST",
                    "LGST",
                    "UNIT PRICE",
                    "DISCOUNT",
                    "REQUIRED DATE (LI)",
                    "EXTENSION",
                    "SHIP DATE",
                    "SHIP COMPLETE",
                    "PO NUMBER",
                    "PROMISE DATE",
                    "OLD PROMISE DATE",
                    "PO LINE NUM",
                    "SUPPLIER NUM",
                    "PURCHASE DATE",
                    "WIK QTY",
                    "WIP QTY",
                    "WIT QTY",
                    "CUSTOMER NAME",
                    "CUSTOMER ADDRESS 1",
                    "CUSTOMER ADDRESS 2",
                    "CUSTOMER CITY",
                    "CUSTOMER STATE",
                    "TRACK ID",
                    "PALLET",
                    "BOX",
                    "QTY",
                    "SUSPENSION TYPE",
                    "COST",
                    "EXT COST",
                    "MARGIN $",
                    "EXT MARGIN $",
                    "QUOTED TO"
                };
                #endregion
            else
                return false;

            //Verify the report type by checking the column headers
            if (TotalCols == ColHeaders.Length)
            {
                for (int i = 0; i < ColHeaders.Length; i++)
                {
                    if (ReportHeaders.Cells[2, i + 1].Value.ToString().Trim() != ColHeaders[i])
                        return false;
                }
            }
            else
                return false;

            return true;
        }
        #endregion

        #region 473
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

        public bool Is473()
        {
            #region Column Headers
            string[] RepHeaders = new string[51]
            {
                "BRANCH",                "ERROR",                "PO NUMBER",                "PO TYPE",                "SHIPPING INSTRUCTIONS 1",                "SHIPPING INSTRUCTIONS 2",                " SUPPLIER",                "SHIP TO",                "DS ORDER",                "PO DATE",                "PO STATUS",                "REQUESTED",                "ACKNOWLEDGE",                "TERMS CODE",                "TERMS DAYS",                "REFERENCE",                "DISC.%",                "FOB",                "BOL",                "SHIPPING TERMS",                "LINE",                "T",                "SIM",                "DESCRIPTION",                "UOM",                "PROMISED",                "QTY ORD",                "QTY REC",                "QTY INV",                "LAST REC",                "PRICE",                "EXTENSION",                "EST",                "ORDER",                "LINE",                "SUPPLIER NAME",                "ADDRESS LINE1",                "ADDRESS LINE2",                "CITY",                "ST",                "ZIP",                "SHIP TO NAME",                "SHIP ADDR LN1",                "SHIP ADDR LN2",                "SHIP CITY",                "SHIP STATE",                "SHIP ZIP",                "NEGNO",                " COSTTYPE",                "COSTDESC",
                "                                                                                                                                                                                                                                                                                                                                                      "
            };
            #endregion

            int TotalCols = ActiveSheet.UsedRange.Columns.Count;
            Excel.Range ReportHeaders = ActiveSheet.Range[ActiveSheet.Cells[2, 1], ActiveSheet.Cells[2, TotalCols]];

            if (ReportHeaders.Columns.Count == RepHeaders.Length)
                for (int i = 0; i < RepHeaders.Length; i++)
                {
                    if (ReportHeaders.Cells[1, i + 1].Value != RepHeaders[i])
                        return false;
                }
            else
                return false;

            return true;
        }
        #endregion

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

        #region GAPS
        public void SaveGAPS()
        {
            DateTime dt = DateTime.Now;
            string Branch = (ActiveSheet.Range["A2"].Value).ToString();
            string sPath = Properties.Settings.Default.PathGAPS + Branch + @" Gaps Download\" + String.Format("{0:yyyy}", dt) + @"\";
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
                "Wdc_rt_qty"
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
        #endregion

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


        public void SaveIROpenOrders()
        {
            string FilePath = @"\\7938-HP02\Shared\IR order entry\IR macro for all plant order entry\IR Open Purchase Orders\";
            string FileName = "Open POs " + Today() + ".xlsx";

            SaveActiveBook(FilePath, FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
        }

        /// <summary>
        /// Finds the specified column and returns the column number.
        /// If no column is found then 0 is returned.
        /// </summary>
        /// <param name="row">The row containg column headers</param>
        /// <param name="text">The column header to search for</param>
        /// <returns></returns>
        private int FindColumnHeader(int row, string text, Excel.Worksheet WS = null)
        {
            if (WS == null)
                WS = ActiveSheet;

            for (int col = 1; col < WS.UsedRange.Columns.Count; col++)
            {
                if (WS.Cells[row, col].Value != null)
                {
                    if ((WS.Cells[row, col].Value).ToString() == text)
                        return col;
                }
            }

            return 0;
        }

        /// <summary>
        /// Save the active workbook. If the file path does not exist it will be created.
        /// </summary>
        /// <param name="Path">Complete path to save location</param>
        /// <param name="FileName">File name and extension</param>
        /// <param name="FileFormat">Excel workbook save type</param>
        private void SaveActiveBook(string Path, string FileName, Excel.XlFileFormat FileFormat)
        {
            try
            {
                if (!Directory.Exists(Path))
                    Directory.CreateDirectory(Path);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "SaveActiveBook - Directory Error");
                return;
            }

            try
            {
                if (Directory.Exists(Path))
                {
                    bool PrevDispAlert = Application.DisplayAlerts;
                    Application.DisplayAlerts = false;
                    ActiveWorkbook.SaveAs(Path + FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    Application.DisplayAlerts = PrevDispAlert;
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                //If error is not due to user canceled save display the error 
                if (e.Message.ToLower() != "exception from hresult: 0x800a03ec")
                    MessageBox.Show(e.Message, "SaveActiveBook - COM Error");
                return;
            }
        }

        private bool IsIROOR()
        {
            #region IR_OOR_Headers
            string[] IR_OOR_Headers = new string[18]
            {
                "Supplier Code",
                "Supplier Name",
                "Location Name",
                "PO Number",
                "Line Number",
                "PO Releases",
                "IR Part Number",
                "IR Part Description",
                "Supplier Part Number",
                "Order Date",
                "Quantity Ordered",
                "Quantity Received",
                "Open Quantity",
                "Performance Date",
                "Actual Due Date",
                "Currency Code",
                "PO Price",
                "Extended PO Price"
            };
            #endregion

            int TotalCols = ActiveSheet.UsedRange.Columns.Count;
            Excel.Range ReportHeaders = ActiveSheet.Range[ActiveSheet.Cells[4, 1], ActiveSheet.Cells[4, TotalCols]];

            if (ReportHeaders.Columns.Count == IR_OOR_Headers.Length)
                for (int i = 0; i < IR_OOR_Headers.Length; i++)
                {
                    if (ReportHeaders.Cells[1, i + 1].Value != IR_OOR_Headers[i])
                        return false;
                }
            else
                return false;

            return true;
        }

        /// <summary>
        /// Gets today plus X number of days and converts it into ISO 8601 compliant string.
        /// </summary>
        /// <param name="Days"></param>
        /// <returns>Todays date in yyyy-MM-dd format</returns>
        private string Today(int Days = 0)
        {
            DateTime dt = DateTime.Now;
            string date = String.Format("{0:yyyy-MM-dd}", dt.AddDays(Days));
            return date;
        }

        public string GetString(Excel.Range value)
        {
            string Result;

            if (value.Value == null)
                Result = String.Empty;
            else
                Result = (value.Value).ToString();

            return Result;
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

        void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            ActiveWorkbook = null;
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
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

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
