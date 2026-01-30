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
// Sat Jan 24 02:00:01 PM UTC 2026
// Sat Jan 24 04:00:01 PM UTC 2026
// Sun Jan 25 10:00:01 AM UTC 2026
// Sun Jan 25 02:00:01 PM UTC 2026
// Sun Jan 25 04:00:01 PM UTC 2026
// Mon Jan 26 10:00:01 AM UTC 2026
// Mon Jan 26 02:00:01 PM UTC 2026
// Mon Jan 26 04:00:01 PM UTC 2026
// Tue Jan 27 10:00:01 AM UTC 2026
// Tue Jan 27 02:00:01 PM UTC 2026
// Tue Jan 27 04:00:01 PM UTC 2026
// Wed Jan 28 10:00:01 AM UTC 2026
// Wed Jan 28 02:00:01 PM UTC 2026
// Wed Jan 28 04:00:01 PM UTC 2026
// Thu Jan 29 10:00:01 AM UTC 2026
// Thu Jan 29 02:00:01 PM UTC 2026
// Thu Jan 29 04:00:01 PM UTC 2026
// Fri Jan 30 10:00:01 AM UTC 2026
// Fri Jan 30 02:00:01 PM UTC 2026
// Fri Jan 30 04:00:01 PM UTC 2026
