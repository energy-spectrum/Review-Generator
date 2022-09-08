using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace CourseWorkOfProgrammingTechnology
{
    class GeneratingReviews
    {
        private GeneralData _generalData;
        private Criticism _advantages;
        private Criticism _disadvantages;
        private List<Student> _students;
        private Random _rand;

        private string _filePathToTemplate;
        private string _filePathToBaseSignature;

        public GeneratingReviews(GeneralData generalData, Criticism advantages, Criticism disadvantages, List<Student> students)
        {
            this._generalData = generalData;
            this._advantages = advantages;
            this._disadvantages = disadvantages;
            this._students = students;
            this._rand = new Random();

            string assmblypath = Assembly.GetEntryAssembly().Location;
            string appPath = Path.GetDirectoryName(assmblypath);

            this._filePathToTemplate = appPath + @"\Properties\Rewio.doc";
            this._filePathToBaseSignature = appPath + @"\Properties\BaseSignature.jpg";
        }

        public void CreateReviews(string folderPath)
        {
            Directory.CreateDirectory(folderPath + "\\pdf");
            Directory.CreateDirectory(folderPath + "\\word");

            WordAPI.Application wordApp = null;
            try
            {
                // Create object of Word
                wordApp = new WordAPI.Application();

                for (int idxStudent = 0; idxStudent < _students.Count; idxStudent++)
                {
                    CreateReview(wordApp, _students[idxStudent], folderPath);
                }
            }
            finally
            {
                // Quit Word.exe
                wordApp.Quit();
                wordApp = null;
            }
        }

        private void CreateReview(WordAPI.Application wordApp, Student student, string folderPath)
        {
            WordAPI.Document wordDoc = null;
            try
            {
                // Add a new document
                wordDoc = wordApp.AddDocument(_filePathToTemplate);

                wordDoc.Replace("faculty", _generalData.Faculty);
                if (_generalData.Type == "Работа")
                {
                    wordDoc.Replace("курсовой проект", "курсовую работу");
                    wordDoc.Replace("курсового проекта", "курсовой работы");
                }
                wordDoc.Replace("Brattsev", student.name);
                wordDoc.Replace("__X__", "__" + _generalData.Course + "__");
                wordDoc.Replace("directionOfTraining", _generalData.DirectionOfTraining);
                wordDoc.Replace("directivity", _generalData.Directivity);
                wordDoc.Replace("topicWork", student.topicWork);
                wordDoc.Replace("academicTitleAndPosition", _generalData.AcademicTitleAndPosition);
                wordDoc.Replace("ФИО_Преподавателя", _generalData.Teacher);
                wordDoc.Replace("rating", student.ConformityRating);

                List<Comment> lAdvantages = ChooseAdvantages(student.rating);
                for (int i = 0; i < lAdvantages.Count; i++)
                {
                    string next = " " + "nextAdvantage";
                    if (i == lAdvantages.Count - 1) { next = string.Empty; }
                    wordDoc.Replace("nextAdvantage", lAdvantages[i].text + next);
                }

                HashSet<int> usedColors = new HashSet<int>();
                for (int i = 0; i < lAdvantages.Count; i++)
                {
                    if (Comment.defaultColor != lAdvantages[i].color)
                        usedColors.Add(lAdvantages[i].color);
                }

                List<Comment> lDisadvantages = ChooseDisadvantages(student.rating, usedColors);
                for (int i = 0; i < lDisadvantages.Count; i++)
                {
                    string next = " " + "nextDisadvantage";
                    if (i == lDisadvantages.Count - 1) { next = string.Empty; }
                    wordDoc.Replace("nextDisadvantage", lDisadvantages[i].text + next);
                }
                wordDoc.Replace("5", student.rating.ToString());
                wordDoc.Replace("И.О. ФамилияРецензента", _generalData.InitialsAndSurnameTeacher);
                wordDoc.Replace("day", _generalData.Date.day);
                wordDoc.Replace("month", _generalData.Date.month);
                wordDoc.Replace("year", _generalData.Date.year);

                ///////////////////////////////////////////////////////////////////////
                WordAPI.Picture signature = null;
                if (_generalData.PathToTeacherSignature != null || _generalData.PathToTeacherSignature != string.Empty)
                {
                    WordAPI.Picture basePicture = wordDoc.AddPicture(mark: "Signature", _filePathToBaseSignature);
                    basePicture.Top += 15;
                    signature = wordDoc.AddPicture("NS", _generalData.PathToTeacherSignature);
                    signature.Top = basePicture.Top;
                    signature.Left = basePicture.Left;
                    basePicture.Delate();
                }
                else
                {
                    wordDoc.Replace("Signature", "");
                }
                //////////////////////////////////////////////////////////////////////
                string fileNameWithoutExtension = @folderPath + "\\pdf" + "\\Рецензия_" + student.lastName;
                wordDoc.SaveAs(fileNameWithoutExtension, extension: ".pdf");

                // Delete the signature, because there should be no signature in the Word file
                signature.Delate();
                //////////////////////////////////////////////////////////////////////
                fileNameWithoutExtension = @folderPath + "\\word" + "\\Рецензия_" + student.lastName;
                // Save the document
                wordDoc.SaveAs(fileNameWithoutExtension, extension: ".docx");
            }
            finally
            {
                wordDoc.Close();
                wordDoc = null;
            }
        }
        private List<Comment> ChooseAdvantages(int rating)
        {
            List<Comment> lAdvantages = new List<Comment>();

            (int, int) range = _advantages.numCommentsForRatings[rating];
            int numComments = _rand.Next(range.Item1, range.Item2 + 1);
            numComments = Math.Min(numComments, _advantages[rating].Count);

            // Shuffling comments for rating evaluation
            Shuffle(_advantages[rating]);
            // Since the comments are shuffled, we choose a random comment each time
            for (int i = 0; i < numComments; i++)
            {
                lAdvantages.Add(_advantages[rating][i]);
            }

            return lAdvantages;
        }

        private List<Comment> ChooseDisadvantages(int rating, HashSet<int> usedColors)
        {
            List<Comment> lDisadvantages = new List<Comment>();

            (int, int) range = _disadvantages.numCommentsForRatings[rating];
            int numComments = _rand.Next(range.Item1, range.Item2 + 1);

            // Shuffling comments for rating evaluation
            Shuffle(_disadvantages[rating]);
            // Since the comments are shuffled, we choose a random comment each time
            for (int i = 0; i < numComments && i < _disadvantages[rating].Count; i++)
            {
                if (usedColors.Contains(_disadvantages[rating][i].color) == false)
                    lDisadvantages.Add(_disadvantages[rating][i]);
                else
                    numComments++;
            }

            return lDisadvantages;
        }

        private void Shuffle(in List<Comment> comments)
        {
            for (int i = comments.Count - 1; i >= 1; i--)
            {
                int j = _rand.Next(i + 1);
                // обменять значения comments[j] и comments[i]
                var temp = comments[j];
                comments[j] = comments[i];
                comments[i] = temp;
            }
        }
    }
}