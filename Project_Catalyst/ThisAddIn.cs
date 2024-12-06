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
// Sat Nov 30 02:00:02 PM CET 2024
// Sat Nov 30 04:00:01 PM CET 2024
// Sun Dec  1 10:00:01 AM CET 2024
// Sun Dec  1 02:00:01 PM CET 2024
// Sun Dec  1 04:00:01 PM CET 2024
// Mon Dec  2 10:00:01 AM CET 2024
// Mon Dec  2 02:00:01 PM CET 2024
// Mon Dec  2 04:00:01 PM CET 2024
// Tue Dec  3 10:00:01 AM CET 2024
// Tue Dec  3 02:00:01 PM CET 2024
// Tue Dec  3 04:00:01 PM CET 2024
// Wed Dec  4 10:00:01 AM CET 2024
// Wed Dec  4 02:00:01 PM CET 2024
// Wed Dec  4 04:00:01 PM CET 2024
// Thu Dec  5 10:00:01 AM CET 2024
// Thu Dec  5 02:00:01 PM CET 2024
// Thu Dec  5 04:00:01 PM CET 2024
// Fri Dec  6 10:00:01 AM CET 2024
