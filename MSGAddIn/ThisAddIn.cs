using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace MSGAddIn
{
    public delegate void ActiveWorksheetChangedEventHeandler(Excel.Worksheet last_wsh, Excel.Worksheet new_wsh);
    
    public partial class ThisAddIn
    {
        private Excel.Worksheet _currentActiveWorkSheet;

        public Excel.Worksheet CurrentActiveWorksheet
        {
            get { return _currentActiveWorkSheet; }
            set {
                var last_wsh = _currentActiveWorkSheet; 
                _currentActiveWorkSheet = value;
                OnActiveWorksheetChanged?.Invoke(last_wsh, _currentActiveWorkSheet);
            }
        }

        public event ActiveWorksheetChangedEventHeandler OnActiveWorksheetChanged;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SheetActivate +=  Application_SheetActivate;
            this.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            if (CurrentActiveWorksheet == null)
                Wb.Worksheets["Начальная"].Activate();
           CurrentActiveWorksheet = Wb.ActiveSheet;
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
             return (Excel.Worksheet)  Application.ActiveSheet;
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
