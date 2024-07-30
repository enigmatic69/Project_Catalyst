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
// Sat Jul 27 02:00:01 PM CEST 2024
// Sat Jul 27 04:00:01 PM CEST 2024
// Sun Jul 28 10:00:01 AM CEST 2024
// Sun Jul 28 02:00:01 PM CEST 2024
// Sun Jul 28 04:00:02 PM CEST 2024
// Mon Jul 29 10:00:01 AM CEST 2024
// Mon Jul 29 02:00:01 PM CEST 2024
// Mon Jul 29 04:00:01 PM CEST 2024
// Tue Jul 30 10:00:01 AM CEST 2024
// Tue Jul 30 02:00:01 PM CEST 2024
// Tue Jul 30 04:00:01 PM CEST 2024
