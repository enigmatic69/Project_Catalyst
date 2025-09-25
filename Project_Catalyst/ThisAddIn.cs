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
// Sat Sep 20 02:00:01 PM CEST 2025
// Sat Sep 20 04:00:01 PM CEST 2025
// Sun Sep 21 10:00:01 AM CEST 2025
// Sun Sep 21 02:00:01 PM CEST 2025
// Sun Sep 21 04:00:01 PM CEST 2025
// Mon Sep 22 10:00:01 AM CEST 2025
// Mon Sep 22 02:00:01 PM CEST 2025
// Mon Sep 22 04:00:01 PM CEST 2025
// Tue Sep 23 10:00:01 AM CEST 2025
// Tue Sep 23 02:00:02 PM CEST 2025
// Tue Sep 23 04:00:01 PM CEST 2025
// Wed Sep 24 10:00:01 AM CEST 2025
// Wed Sep 24 02:00:01 PM CEST 2025
// Wed Sep 24 04:00:01 PM CEST 2025
// Thu Sep 25 10:00:01 AM CEST 2025
