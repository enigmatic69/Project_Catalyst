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
// Sat 08 Jun 2024 02:00:01 PM +08
// Sat 08 Jun 2024 04:00:01 PM +08
// Sun 09 Jun 2024 10:00:01 AM +08
// Sun 09 Jun 2024 02:00:01 PM +08
// Sun 09 Jun 2024 04:00:01 PM +08
// Mon 10 Jun 2024 10:00:01 AM +08
// Mon 10 Jun 2024 02:00:01 PM +08
// Mon 10 Jun 2024 04:00:01 PM +08
// Tue 11 Jun 2024 10:00:01 AM +08
// Tue 11 Jun 2024 02:00:01 PM +08
// Tue 11 Jun 2024 04:00:01 PM +08
// Wed 12 Jun 2024 10:00:02 AM +08
// Wed 12 Jun 2024 02:00:01 PM +08
// Wed 12 Jun 2024 04:00:01 PM +08
// Thu 13 Jun 2024 10:00:01 AM +08
// Thu 13 Jun 2024 02:00:01 PM +08
// Thu 13 Jun 2024 04:00:02 PM +08
// Fri 14 Jun 2024 10:00:01 AM +08
// Fri 14 Jun 2024 02:00:01 PM +08
// Fri 14 Jun 2024 04:00:01 PM +08
