using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace AssetTrackingApplication
{
    public class Excel
    {
        public Excel(string fileName)
        {
            FileName = fileName;
            Application = new Application();
            Workbook = Application.Workbooks.Open(FileName, ReadOnly: false, Editable: true);
            Worksheet = Workbook.Worksheets.Item[1];
        }
        public string FileName { get; set; }
        public Application Application { get; set; }
        public Workbook Workbook { get; set; }
        public Worksheet Worksheet { get; set; }

        public void CloseExcelFile()
        {
            Workbook.Close(true, FileName);
            Application.Quit();
            Marshal.ReleaseComObject(Worksheet);
            Marshal.ReleaseComObject(Workbook);
            Marshal.ReleaseComObject(Application);
        }
    }
}
