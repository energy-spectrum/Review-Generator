using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace CourseWorkOfProgrammingTechnology
{
    class SystemInputData
    {
        private void ShowFileDialog(OpenFileDialog openFileDialog, string filter, out string filePath)
        {
            filePath = string.Empty;
            //Universal path to the desktop
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            openFileDialog.Filter = filter;
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }
        }

        private string _filePath = string.Empty;// = @"C:\Users\batar\Desktop\Шаблон исходных данных к КР (Братцев).xlsx";
        public string FilePath { get { return _filePath; } }
        public void ShowFileDialogForInputDateFile(OpenFileDialog openFileDialog)
        {
            ShowFileDialog(openFileDialog, filter: "excel files (*.xlsx)|*.xlsx", out _filePath);
        }

        private string _pathToTeacherSignature = string.Empty;
        public string PathToTeacherSignature { get { return _pathToTeacherSignature; } }
        public void ShowFileDialogForSignature(OpenFileDialog openFileDialog )
        {
            ShowFileDialog(openFileDialog, filter: "(*.jp2)|*.jp2| (*.jpg)|*.jpg| (*.png)|*.png", out _pathToTeacherSignature);
        }

        private string _folderPath = string.Empty;
        public string FolderPath { get { return _folderPath; } }
        public void ShowFolderDialog(FolderBrowserDialog folderBrowserDialog)
        {
            _folderPath = string.Empty;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                _folderPath = folderBrowserDialog.SelectedPath;
            }
        }

        public bool IsInputCompleted()
        {
            return _filePath != string.Empty && _pathToTeacherSignature != string.Empty && _folderPath != string.Empty;
        }

        private string ConvertExcelCellToString(dynamic excelCell)
        {
            Excel.Range exCel = excelCell as Excel.Range;
            string res = Convert.ToString(exCel.Value);

            Marshal.ReleaseComObject(exCel);

            return res;
        }

        private GeneralData GetGeneralData(Excel.Worksheet generalDataSheet)
        {
            Excel.Range cells = generalDataSheet.Cells;

            GeneralData generalData = new GeneralData();
            int row = 2;
            generalData.Faculty = ConvertExcelCellToString(cells[row++, "B"]);
            if (ConvertExcelCellToString(cells[row++, "B"]) == "+")
            {
                generalData.Type = "Проект";
            }
            if (ConvertExcelCellToString(cells[row++, "B"]) == "+")
            {
                generalData.Type = "Работа";
            }
            generalData.Course = ConvertExcelCellToString(cells[row++, "B"]);
            generalData.DirectionOfTraining = ConvertExcelCellToString(cells[row++, "B"]);
            generalData.Directivity = ConvertExcelCellToString(cells[row++, "B"]);
            generalData.Teacher = ConvertExcelCellToString(cells[row++, "B"]);
            generalData.AcademicTitleAndPosition = ConvertExcelCellToString(cells[row++, "B"]);
            generalData.Group = ConvertExcelCellToString(cells[row++, "B"]);
            generalData.Date = new Date(ConvertExcelCellToString(cells[row++, "B"]));

            generalData.PathToTeacherSignature = _pathToTeacherSignature;

            return generalData;
        }
        private Criticism GetCriticism(Excel.Worksheet commentsSheet)
        {
            Criticism criticism = new Criticism();

            int row = 2;
            while (ConvertExcelCellToString(commentsSheet.Cells[row, "A"]) != null)
            {
                Comment comment = new Comment(ConvertExcelCellToString(commentsSheet.Cells[row, "A"]), (commentsSheet.Cells[row, "A"] as Excel.Range).Interior.ColorIndex);
                if (ConvertExcelCellToString(commentsSheet.Cells[row, "B"]) == "+") { criticism[3].Add(comment); }
                if (ConvertExcelCellToString(commentsSheet.Cells[row, "C"]) == "+") { criticism[4].Add(comment); }
                if (ConvertExcelCellToString(commentsSheet.Cells[row, "D"]) == "+") { criticism[5].Add(comment); }

                row++;
            }

            (int, int) GetRange(string sRange)
            {
                (int, int) range = (0, 0);
                string[] aRange = sRange.Split('-');
                range.Item1 = int.Parse(aRange[0]);

                if (aRange.Length == 1)
                {
                    range.Item2 = int.Parse(aRange[0]);
                }
                else if (aRange.Length == 2)
                {
                    range.Item2 = int.Parse(aRange[1]);
                }

                return range;
            }
            
            criticism.numCommentsForRatings[5] = GetRange(ConvertExcelCellToString(commentsSheet.Cells[2, "G"]));
            criticism.numCommentsForRatings[4] = GetRange(ConvertExcelCellToString(commentsSheet.Cells[3, "G"]));
            criticism.numCommentsForRatings[3] = GetRange(ConvertExcelCellToString(commentsSheet.Cells[4, "G"]));

            return criticism;
        }
        private List<Student> GetStudents(Excel.Worksheet listOfStudentsSheet)
        {
            List<Student> students = new List<Student>();
            int row = 2;
            while (ConvertExcelCellToString(listOfStudentsSheet.Cells[row, "A"]) != null)
            {
                string name = ConvertExcelCellToString(listOfStudentsSheet.Cells[row, "A"]);
                string nameInGenitiveCase = ConvertExcelCellToString(listOfStudentsSheet.Cells[row, "E"]);
                string topicWork = ConvertExcelCellToString(listOfStudentsSheet.Cells[row, "B"]);
                int rating = int.Parse(ConvertExcelCellToString(listOfStudentsSheet.Cells[row, "C"]));
                students.Add(new Student(name, nameInGenitiveCase, topicWork, rating));
                row++;
            }

            return students;
        }

        public (GeneralData, Criticism, Criticism, List<Student>) GetInputData()
        {
            if (_filePath == string.Empty) { return (null, null, null, null); }

            Object oMissing = System.Reflection.Missing.Value;

            GeneralData generalData = null;
            Criticism advantages = null;
            Criticism disadvantages = null;
            List<Student> students = null;

            Comment.defaultColor = (int)Excel.XlColorIndex.xlColorIndexNone;

            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook inputDataBook = null;
            Excel.Sheets sheets = null;
            
            try
            {
                // Create object of Excel.
                excelApp = new Excel.Application();
                workbooks = excelApp.Workbooks;
                // Open the workbook for read-only.
                inputDataBook = workbooks.Open(
                    _filePath,
                    oMissing, true, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing,
                    oMissing, oMissing);
                sheets = inputDataBook.Sheets;

                Excel.Worksheet generalDataSheet = (Excel.Worksheet)sheets[1];
                generalData = GetGeneralData(generalDataSheet);
                Marshal.ReleaseComObject(generalDataSheet);
                generalDataSheet = null;

                Excel.Worksheet advantagesSheet = (Excel.Worksheet)sheets[2];
                advantages = GetCriticism(advantagesSheet);
                Marshal.ReleaseComObject(advantagesSheet);
                advantagesSheet = null;

                Excel.Worksheet disadvantagesSheet = (Excel.Worksheet)sheets[3];
                disadvantages = GetCriticism(disadvantagesSheet);
                Marshal.ReleaseComObject(disadvantagesSheet);
                disadvantagesSheet = null;

                Excel.Worksheet studentsSheet = (Excel.Worksheet)sheets[4];
                students = GetStudents(studentsSheet);
                Marshal.ReleaseComObject(studentsSheet);
                studentsSheet = null;
            }
            finally
            {
                inputDataBook.Close(false, oMissing, oMissing);
                workbooks.Close();
                excelApp.Quit();

                Marshal.ReleaseComObject(sheets);
                Marshal.ReleaseComObject(inputDataBook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(excelApp);

                sheets = null;
                inputDataBook = null;
                workbooks = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return (generalData, advantages, disadvantages, students);
        }//конец функции
    }
}