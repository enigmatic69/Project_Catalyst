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
// Sat Apr 11 02:00:01 PM UTC 2026
// Sat Apr 11 04:00:01 PM UTC 2026
// Sun Apr 12 10:00:01 AM UTC 2026
// Sun Apr 12 02:00:01 PM UTC 2026
// Sun Apr 12 04:00:01 PM UTC 2026
// Mon Apr 13 10:00:01 AM UTC 2026
// Mon Apr 13 02:00:01 PM UTC 2026
// Mon Apr 13 04:00:01 PM UTC 2026
// Tue Apr 14 10:00:01 AM UTC 2026
// Tue Apr 14 02:00:02 PM UTC 2026
// Tue Apr 14 04:00:01 PM UTC 2026
// Wed Apr 15 10:00:01 AM UTC 2026
// Wed Apr 15 02:00:01 PM UTC 2026
// Wed Apr 15 04:00:01 PM UTC 2026
// Thu Apr 16 10:00:01 AM UTC 2026
// Thu Apr 16 02:00:01 PM UTC 2026
// Thu Apr 16 04:00:02 PM UTC 2026
// Fri Apr 17 10:00:01 AM UTC 2026
