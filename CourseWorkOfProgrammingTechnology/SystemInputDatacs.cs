using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace CourseWorkOfProgrammingTechnology
{
    class SystemInputData
    {
        private string _filePath = @"C:\Users\batar\Desktop\Курсавая работа\Шаблон исходных данных к КР (Братцев).xlsx";// = @"C:\Users\batar\Desktop\Шаблон исходных данных к КР (Братцев).xlsx";
        public string FilePath { get { return _filePath; } }

        private string _pathToTeacherSignature = @"C:\Users\batar\Desktop\Курсавая работа\ПодписьРецензента.jpg";
        public string PathToTeacherSignature { get { return _pathToTeacherSignature; } }

        private string _folderPath = string.Empty;
        public string FolderPath { get { return _folderPath; } }

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
        public void ShowFileDialogForInputDateFile(OpenFileDialog openFileDialog)
        {
            ShowFileDialog(openFileDialog, filter: "excel files (*.xlsx)|*.xlsx", out _filePath);
        }
        public void ShowFileDialogForSignature(OpenFileDialog openFileDialog)
        {
            ShowFileDialog(openFileDialog, filter: "(*.jp2)|*.jp2| (*.jpg)|*.jpg| (*.png)|*.png", out _pathToTeacherSignature);
        }
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

        public (GeneralData, Criticism, Criticism, List<Student>) GetInputData()
        {
            if (_filePath == string.Empty) { return (null, null, null, null); }

            GeneralData generalData = null;
            Criticism advantages = null;
            Criticism disadvantages = null;
            List<Student> students = null;

            ExcelAPI.Application excelAplication = null;

            try
            {
                excelAplication = new ExcelAPI.Application(_filePath);

                ExcelAPI.Cells cellsGeneralDataSheet = excelAplication.GetCellsWorkshit(1);
                generalData = GetGeneralData(cellsGeneralDataSheet);

                ExcelAPI.Cells cellsAdvantagesSheet = excelAplication.GetCellsWorkshit(2);
                advantages = GetCriticism(cellsAdvantagesSheet);

                ExcelAPI.Cells cellsDisadvantagesSheet = excelAplication.GetCellsWorkshit(3);
                disadvantages = GetCriticism(cellsDisadvantagesSheet);

                ExcelAPI.Cells cellsStudentsSheet = excelAplication.GetCellsWorkshit(4);
                students = GetStudents(cellsStudentsSheet);
            }
            catch
            {
                Debug.WriteLine("Error ExcelAPI ban");
            }
            finally
            {
                excelAplication?.Quit();
            }

            return (generalData, advantages, disadvantages, students);
        }

        private GeneralData GetGeneralData(ExcelAPI.Cells cellsGeneralDataSheet)
        {
            GeneralData generalData = new GeneralData();
            int row = 2;
            generalData.Faculty = cellsGeneralDataSheet[row++, "B"];
            if (cellsGeneralDataSheet[row++, "B"] == "+")
            {
                generalData.Type = "Проект";
            }
            if (cellsGeneralDataSheet[row++, "B"] == "+")
            {
                generalData.Type = "Работа";
            }
            generalData.Course = cellsGeneralDataSheet[row++, "B"];
            generalData.DirectionOfTraining = cellsGeneralDataSheet[row++, "B"];
            generalData.Directivity = cellsGeneralDataSheet[row++, "B"];
            generalData.Teacher = cellsGeneralDataSheet[row++, "B"];
            generalData.AcademicTitleAndPosition = cellsGeneralDataSheet[row++, "B"];
            generalData.Group = cellsGeneralDataSheet[row++, "B"];
            generalData.Date = new Date(cellsGeneralDataSheet[row++, "B"]);
            generalData.PathToTeacherSignature = _pathToTeacherSignature;

            return generalData;
        }
        private Criticism GetCriticism(ExcelAPI.Cells cellsCommentsSheet)
        {
            Criticism criticism = new Criticism();

            int row = 2;
            while (cellsCommentsSheet[row, "A"] != null)
            {
                Comment comment = new Comment(cellsCommentsSheet[row, "A"], cellsCommentsSheet.GetCellColorIndex(row, "A"));
                if (cellsCommentsSheet[row, "B"] == "+") { criticism[3].Add(comment); }
                if (cellsCommentsSheet[row, "C"] == "+") { criticism[4].Add(comment); }
                if (cellsCommentsSheet[row, "D"] == "+") { criticism[5].Add(comment); }

                row++;
            }

            criticism.numCommentsForRatings[5] = GetRange(cellsCommentsSheet[2, "G"]);
            criticism.numCommentsForRatings[4] = GetRange(cellsCommentsSheet[3, "G"]);
            criticism.numCommentsForRatings[3] = GetRange(cellsCommentsSheet[4, "G"]);

            return criticism;

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
        }
        private List<Student> GetStudents(ExcelAPI.Cells cellsListOfStudentsSheet)
        {
            List<Student> students = new List<Student>();
            int row = 2;
            while (cellsListOfStudentsSheet[row, "A"] != null)
            {
                string name = cellsListOfStudentsSheet[row, "A"];
                string nameInGenitiveCase = cellsListOfStudentsSheet[row, "E"];
                string topicWork = cellsListOfStudentsSheet[row, "B"];
                int rating = int.Parse(cellsListOfStudentsSheet[row, "C"]);
                students.Add(new Student(name, nameInGenitiveCase, topicWork, rating));
                row++;
            }

            return students;
        }
    }
}