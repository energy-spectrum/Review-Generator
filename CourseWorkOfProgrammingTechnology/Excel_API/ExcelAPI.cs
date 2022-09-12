using System;
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

        public event Action OnQuit;

        public Application()
        {
            RunExcelApplication();
        }

        private void RunExcelApplication()
        {
            try
            {
                TryRunExcelApplication();
            }
            catch
            {
                Quit();
                throw new Exception("Error run Excel application");
            }
        }

        private void TryRunExcelApplication()
        {
            _excelApp = new Excel.Application();
            _workbooks = _excelApp.Workbooks;
        }

        public void OpenExcelFile(string filePath)
        {
            try
            {
                TryOpenExcelFile(filePath);
            }
            catch
            {
                Quit();
                throw new Exception("Error open file: " + filePath);
            }
        }

        private void TryOpenExcelFile(string filePath)
        {
            _inputDataBook = _workbooks.Open(Filename: filePath, ReadOnly: true);
            _sheets = _inputDataBook.Sheets;
        }

        public void Quit()
        {
            OnQuit?.Invoke();
            CloseExcelApplication();
            ReleaseAllComObjects();
            ClearMemory();
        }

        private void CloseExcelApplication()
        {
            _inputDataBook?.Close(SaveChanges: false);
            _workbooks?.Close();
            _excelApp?.Quit();
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

        private void ClearMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public Cells GetCellsWorkshit(int index)
        {
            Cells cells = new Cells((Excel.Worksheet)_sheets[index]);
            OnQuit += cells.ReleaseAllComObjects;

            return cells;
        }
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

        public void ReleaseAllComObjects()
        {
            Marshal.ReleaseComObject(_worksheet);
            _worksheet = null;
        }
    }
}
