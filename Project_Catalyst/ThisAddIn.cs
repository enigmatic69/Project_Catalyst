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
// Sat Dec 28 02:00:01 PM CET 2024
// Sat Dec 28 04:00:01 PM CET 2024
// Sun Dec 29 10:00:01 AM CET 2024
// Sun Dec 29 02:00:01 PM CET 2024
// Sun Dec 29 04:00:01 PM CET 2024
// Mon Dec 30 10:00:01 AM CET 2024
// Mon Dec 30 02:00:01 PM CET 2024
// Mon Dec 30 04:00:01 PM CET 2024
// Tue Dec 31 10:00:01 AM CET 2024
// Tue Dec 31 02:00:01 PM CET 2024
// Tue Dec 31 04:00:01 PM CET 2024
// Wed Jan  1 10:00:01 AM CET 2025
// Wed Jan  1 02:00:01 PM CET 2025
// Wed Jan  1 04:00:01 PM CET 2025
// Thu Jan  2 10:00:01 AM CET 2025
// Thu Jan  2 02:00:01 PM CET 2025
// Thu Jan  2 04:00:01 PM CET 2025
// Fri Jan  3 10:00:01 AM CET 2025
// Fri Jan  3 02:00:02 PM CET 2025
// Fri Jan  3 04:00:01 PM CET 2025
