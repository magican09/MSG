using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSGAddIn
{
    public delegate void ActiveWorksheetChangedEventHeandler(Excel.Worksheet last_wsh, Excel.Worksheet new_wsh);
    public delegate void ActiveWorkbookChangedEventHeandler(Excel.Workbook last_wbk, Excel.Workbook new_wbk);

    public partial class ThisAddIn
    {
        public event ActiveWorksheetChangedEventHeandler OnActiveWorksheetChanged;
        public event ActiveWorkbookChangedEventHeandler OnActiveWorkbookChanged;
        private Excel.Worksheet _currentActiveWorkSheet;

        public Excel.Worksheet CurrentActiveWorksheet
        {
            get { return _currentActiveWorkSheet; }
            set
            {
                var last_wsh = _currentActiveWorkSheet;
                _currentActiveWorkSheet = value;
                OnActiveWorksheetChanged?.Invoke(last_wsh, _currentActiveWorkSheet);
            }
        }

        private Excel.Workbook _currentActivWorkbook;
        public Excel.Workbook CurrentActivWorkbook
        {
            get { return _currentActivWorkbook; }
            set
            {
                var last_wbk = _currentActivWorkbook;
                _currentActivWorkbook = value;
                OnActiveWorkbookChanged?.Invoke(last_wbk, _currentActivWorkbook);
            }
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SheetActivate += Application_SheetActivate;
            this.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            //if (CurrentActiveWorksheet == null)
            //    Wb.Worksheets["Начальная"].Activate();
            CurrentActiveWorksheet = Wb.ActiveSheet;
            CurrentActivWorkbook = Wb;
        }

        private void Application_SheetActivate(object Sh)
        {
            CurrentActiveWorksheet = (Excel.Worksheet)Sh;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

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
