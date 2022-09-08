using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAPI
{
    class Application
    {
        private Excel.Application _excelApp = null;
        private Excel.Workbooks _workbooks = null;
        private Excel.Workbook _inputDataBook = null;
        private Excel.Sheets _sheets = null;

        private Object oMissing = System.Reflection.Missing.Value;

        public event Action OnQuit;

        public Cells GetCellsWorkshit(int index)
        {
            Cells cells = new Cells((Excel.Worksheet)_sheets[index]);
            OnQuit += cells.Quit;
            return cells;
        }

        public Application(string filePath)
        {
            TryOpenExcelFile(filePath);
        }

        private void TryOpenExcelFile(string filePath)
        {
            try
            {
                OpenExcelFile(filePath);
            }
            catch
            {
                Debug.WriteLine("Error open file");
                Quit();
            }
        }

        private void OpenExcelFile(string filePath)
        {
            // Create object of Excel.
            _excelApp = new Excel.Application();
            _workbooks = _excelApp.Workbooks;
            // Open the workbook for read-only.
            _inputDataBook = _workbooks.Open(
                filePath,
                oMissing, true, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing);

            _sheets = _inputDataBook.Sheets;
        }

        public void Quit()
        {
            OnQuit?.Invoke();
            CloseExcelApplication();
            ReleaseAllComObjects();
        }

        private void CloseExcelApplication()
        {
            _inputDataBook.Close(false, oMissing, oMissing);
            _workbooks.Close();
            _excelApp.Quit();
        }

        private void ReleaseAllComObjects()
        {
            Marshal.ReleaseComObject(_sheets);
            _sheets = null;
            Marshal.ReleaseComObject(_inputDataBook);
            _inputDataBook = null;
            Marshal.ReleaseComObject(_workbooks);
            _workbooks = null;
            Marshal.ReleaseComObject(_excelApp);
            _excelApp = null;

        }

        ////private void Clear()
        ////{
        ////    //GC.Collect();
        ////    //GC.WaitForPendingFinalizers();
        ////}
    }

    class Cells
    {
        private Excel.Worksheet _worksheet;
        public Cells(Excel.Worksheet worksheet)
        {
            this._worksheet = worksheet;
        }

        public string this[int numerRow, string nameCol]
        {
            get
            {
                Excel.Range excelCell = _worksheet.Cells[numerRow, nameCol] as Excel.Range;
                string value = ConvertExcelCellToString(excelCell);
                Marshal.ReleaseComObject(excelCell);

                return value;
            }
        }

        private string ConvertExcelCellToString(Excel.Range excelCell)
        {
            string value = Convert.ToString(excelCell.Value);
            return value;
        }

        public int GetCellColorIndex(int numerRow, string nameCol)
        {
            Excel.Range excelCell = _worksheet.Cells[numerRow, nameCol] as Excel.Range;
            int colorIndex = excelCell.Interior.ColorIndex;
            Marshal.ReleaseComObject(excelCell);

            return colorIndex;
        }

        public void Quit()
        {
            Marshal.ReleaseComObject(_worksheet);
            _worksheet = null;
        }
    }
}
