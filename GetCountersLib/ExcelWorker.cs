using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace GetCountersLib
{
    public class ExcelWorker
    {
        private readonly Excel.Application _excelApp;
        private readonly Excel.Workbook _excelWorkbook;
        private readonly Excel.Worksheet _excelWorksheet;
        private int _down = 1;

        public ExcelWorker()
        {
            _excelApp = new Excel.Application();
            _excelWorkbook = _excelApp.Workbooks.Add();
            _excelWorksheet = (Excel.Worksheet) _excelWorkbook.Sheets.Add();

            _excelWorksheet.Name = "Test";

            _excelWorksheet.Cells[_down, 1] = "Time";
            _excelWorksheet.Cells[_down, 2] = "Process";
            _excelWorksheet.Cells[_down, 3] = "Ram";
            _excelWorksheet.Cells[_down, 4] = "Disk";
            _down++;
        }

        public void AddToExcelFile(IList<CountersModel> models)
        {
            foreach (var model in models)
            {
                _excelWorksheet.Cells[_down, 1] = model.DateTime;
                _excelWorksheet.Cells[_down, 2] = model.CpuCounter;
                _excelWorksheet.Cells[_down, 3] = model.RamCounter;
                _excelWorksheet.Cells[_down, 4] = model.DiskCounter;
                _down++;
            }
        }
        public void CloseExcelFile()
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelWorksheet);
            _excelApp.ActiveWorkbook.SaveAs($@"D:\abs.xls", Excel.XlFileFormat.xlWorkbookNormal);

            _excelWorkbook.Close();
            _excelApp.Quit();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
