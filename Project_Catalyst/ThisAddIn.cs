﻿using System;
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
// Sat Jan  4 02:00:01 PM CET 2025
// Sat Jan  4 04:00:01 PM CET 2025
// Sun Jan  5 10:00:01 AM CET 2025
// Sun Jan  5 02:00:01 PM CET 2025
// Sun Jan  5 04:00:01 PM CET 2025
// Mon Jan  6 10:00:01 AM CET 2025
// Mon Jan  6 02:00:01 PM CET 2025
// Mon Jan  6 04:00:01 PM CET 2025
// Tue Jan  7 10:00:01 AM CET 2025
// Tue Jan  7 02:00:01 PM CET 2025
// Tue Jan  7 04:00:01 PM CET 2025
// Wed Jan  8 10:00:01 AM CET 2025
// Wed Jan  8 02:00:01 PM CET 2025
// Wed Jan  8 04:00:01 PM CET 2025
// Thu Jan  9 10:00:01 AM CET 2025
// Thu Jan  9 02:00:01 PM CET 2025
// Thu Jan  9 04:00:01 PM CET 2025
// Fri Jan 10 10:00:01 AM CET 2025
