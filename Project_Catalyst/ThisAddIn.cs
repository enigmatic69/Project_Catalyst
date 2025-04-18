using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Project_Catalyst
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = DateTime.Now.ToString();
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }
}
// Sat Apr 12 02:00:01 PM CEST 2025
// Sat Apr 12 04:00:01 PM CEST 2025
// Sun Apr 13 10:00:01 AM CEST 2025
// Sun Apr 13 02:00:01 PM CEST 2025
// Sun Apr 13 04:00:01 PM CEST 2025
// Mon Apr 14 10:00:01 AM CEST 2025
// Mon Apr 14 02:00:01 PM CEST 2025
// Mon Apr 14 04:00:01 PM CEST 2025
// Tue Apr 15 10:00:02 AM CEST 2025
// Tue Apr 15 02:00:01 PM CEST 2025
// Tue Apr 15 04:00:01 PM CEST 2025
// Wed Apr 16 10:00:01 AM CEST 2025
// Wed Apr 16 02:00:01 PM CEST 2025
// Wed Apr 16 04:00:01 PM CEST 2025
// Thu Apr 17 10:00:01 AM CEST 2025
// Thu Apr 17 02:00:01 PM CEST 2025
// Thu Apr 17 04:00:01 PM CEST 2025
// Fri Apr 18 10:00:01 AM CEST 2025
