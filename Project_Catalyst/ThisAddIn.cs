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
// Sat Feb  8 02:00:01 PM CET 2025
// Sat Feb  8 04:00:01 PM CET 2025
// Sun Feb  9 10:00:01 AM CET 2025
// Sun Feb  9 02:00:01 PM CET 2025
// Sun Feb  9 04:00:01 PM CET 2025
// Mon Feb 10 10:00:01 AM CET 2025
// Mon Feb 10 02:00:02 PM CET 2025
// Mon Feb 10 04:00:01 PM CET 2025
// Tue Feb 11 10:00:02 AM CET 2025
// Tue Feb 11 02:00:01 PM CET 2025
// Tue Feb 11 04:00:01 PM CET 2025
// Wed Feb 12 10:00:01 AM CET 2025
// Wed Feb 12 02:00:01 PM CET 2025
// Wed Feb 12 04:00:01 PM CET 2025
// Thu Feb 13 10:00:01 AM CET 2025
// Thu Feb 13 02:00:01 PM CET 2025
// Thu Feb 13 04:00:01 PM CET 2025
// Fri Feb 14 10:00:01 AM CET 2025
// Fri Feb 14 02:00:01 PM CET 2025
// Fri Feb 14 04:00:01 PM CET 2025
// Sat Feb 15 10:00:01 AM CET 2025
