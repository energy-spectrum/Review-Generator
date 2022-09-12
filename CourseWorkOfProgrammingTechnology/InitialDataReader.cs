using System;
using System.Collections.Generic;

namespace CourseWorkOfProgrammingTechnology
{
    class InitialDataReader
    {
        public InitialData GetInitialDataFromFile(string pathToExcelFile)
        {
            ExcelAPI.Application excelApp = null;
            InitialData initialData;
            try
            {
                excelApp = new ExcelAPI.Application();
                initialData = TryGetInitialDataFromFile(excelApp, pathToExcelFile);
            }
            finally
            {
                excelApp?.Quit();
            }

            return initialData;
        }

        private InitialData TryGetInitialDataFromFile(ExcelAPI.Application excelApp, string pathToExcelFile)
        {
            excelApp.OpenExcelFile(pathToExcelFile);

            const int idxWorkshitWithGeneralData = 1;
            const int idxWorkshitWithAdvantages = 2;
            const int idxWorkshitWithDisadvantages = 3;
            const int idxWorkshitWithStudents = 4;

            ExcelAPI.Cells cellsGeneralDataSheet = excelApp.GetCellsWorkshit(idxWorkshitWithGeneralData);
            GeneralData generalData = GetGeneralData(cellsGeneralDataSheet);

            ExcelAPI.Cells cellsAdvantagesSheet = excelApp.GetCellsWorkshit(idxWorkshitWithAdvantages);
            Criticism advantages = GetCriticism(cellsAdvantagesSheet);

            ExcelAPI.Cells cellsDisadvantagesSheet = excelApp.GetCellsWorkshit(idxWorkshitWithDisadvantages);
            Criticism disadvantages = GetCriticism(cellsDisadvantagesSheet);

            ExcelAPI.Cells cellsStudentsSheet = excelApp.GetCellsWorkshit(idxWorkshitWithStudents);
            List<Student> students = GetStudents(cellsStudentsSheet);

            InitialData inputData = new InitialData(generalData, advantages, disadvantages, students);

            return inputData;
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
                string topicWork = cellsListOfStudentsSheet[row, "B"];
                int rating = int.Parse(cellsListOfStudentsSheet[row, "C"]);
                students.Add(new Student(name, topicWork, rating));
                row++;
            }

            return students;
        }
    }
}