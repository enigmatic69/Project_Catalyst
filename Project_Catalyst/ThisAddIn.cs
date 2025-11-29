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
// Sat Nov 22 02:00:01 PM CET 2025
// Sat Nov 22 04:00:01 PM CET 2025
// Sun Nov 23 10:00:01 AM CET 2025
// Sun Nov 23 02:00:01 PM CET 2025
// Sun Nov 23 04:00:01 PM CET 2025
// Mon Nov 24 10:00:01 AM CET 2025
// Mon Nov 24 02:00:01 PM CET 2025
// Mon Nov 24 04:00:01 PM CET 2025
// Tue Nov 25 10:00:01 AM CET 2025
// Tue Nov 25 02:00:01 PM CET 2025
// Tue Nov 25 04:00:02 PM CET 2025
// Wed Nov 26 10:00:01 AM CET 2025
// Wed Nov 26 02:00:01 PM CET 2025
// Wed Nov 26 04:00:01 PM CET 2025
// Thu Nov 27 10:00:01 AM CET 2025
// Thu Nov 27 02:00:01 PM CET 2025
// Thu Nov 27 04:00:01 PM CET 2025
// Fri Nov 28 10:00:01 AM CET 2025
// Fri Nov 28 02:00:01 PM CET 2025
// Fri Nov 28 04:00:01 PM CET 2025
// Sat Nov 29 10:00:02 AM CET 2025
