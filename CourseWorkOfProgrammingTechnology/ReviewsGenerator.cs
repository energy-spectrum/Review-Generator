using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace CourseWorkOfProgrammingTechnology
{
    class ReviewsGenerator
    {
        private InitialData _initialData;
        private Critic _critic;

        private string _filePathToTemplate;
        private string _filePathToBaseSignature;
        private string _pathToTeacherSignature;
        private string _folderPath;
        public ReviewsGenerator(InitialData initialData, string pathToTeacherSignature, string folderPath)
        {
            this._initialData = initialData;
            this._critic = new Critic(_initialData.advantages, _initialData.disadvantages);

            SetPathsToTemplateAndBaseSignature();
            this._pathToTeacherSignature = pathToTeacherSignature;
            this._folderPath = folderPath;
        }

        private void SetPathsToTemplateAndBaseSignature()
        {
            string assmblypath = Assembly.GetEntryAssembly().Location;
            string appPath = Path.GetDirectoryName(assmblypath);

            this._filePathToTemplate = appPath + @"\Properties\Review.doc";
            this._filePathToBaseSignature = appPath + @"\Properties\BaseSignature.jpg";
        }

        public void CreateReviews()
        {
            List<Student> students = _initialData.students;

            CreateDirectories();

            WordAPI.Application wordApp = null;
            try
            {
                wordApp = new WordAPI.Application();

                for (int idxStudent = 0; idxStudent < students.Count; idxStudent++)
                {
                    CreateReview(wordApp, students[idxStudent]);
                }
            }
            finally
            {
                wordApp?.Quit();
                wordApp = null;
            }
        }

        private void CreateDirectories()
        {
            Directory.CreateDirectory(_folderPath + "\\pdf");
            Directory.CreateDirectory(_folderPath + "\\word");
        }

        private void CreateReview(WordAPI.Application wordApp, Student student)
        {
            WordAPI.Document wordDoc = null;
            try
            {
                wordDoc = wordApp.AddDocument(_filePathToTemplate);
                TryCreateReview(wordDoc, student);
            }
            finally
            {
                wordDoc.Close();
                wordDoc = null;
            }
        }

        private void TryCreateReview(WordAPI.Document wordDoc, Student student)
        {
            GeneralData generalData = _initialData.generalData;

            wordDoc.Replace("[Faculty]", generalData.Faculty);

            if (generalData.Type == "Работа")
            {
                wordDoc.Replace("[курсовой проект]", "курсовую работу");
                wordDoc.Replace("[курсового проекта]", "курсовой работы");
            }
            else
            {
                //Delete []
                wordDoc.Replace("[курсовой проект]", "курсовой проект");
                wordDoc.Replace("[курсового проекта]", "курсового проекта");
            }

            wordDoc.Replace("[LastNameFirstNamePatronymicOfStudentWhose]", student.name);
            wordDoc.Replace("[Course]", generalData.Course);
            wordDoc.Replace("[DirectionOfTraining]", generalData.DirectionOfTraining);
            wordDoc.Replace("[Directivity]", generalData.Directivity);
            wordDoc.Replace("[TopicWork]", student.topicWork);
            wordDoc.Replace("[AcademicTitleAndPosition]", generalData.AcademicTitleAndPosition);
            wordDoc.Replace("[LastNameFirstNamePatronymicOfReviewer]", generalData.Teacher);
            wordDoc.Replace("[RatingInVerbalForm]", student.GetRatingInVerbalForm());

            _critic.AddAdvantagesAndDisadvantages(wordDoc, student);

            wordDoc.Replace("[Rating]", student.rating.ToString());
            wordDoc.Replace("[И.О. ФамилияРецензента]", generalData.InitialsAndSurnameTeacher);
            wordDoc.Replace("[day]", generalData.Date.day);
            wordDoc.Replace("[month]", generalData.Date.month);
            wordDoc.Replace("[year]", generalData.Date.year);

            WordAPI.Picture signature = AddSignature(wordDoc);
            SaveAs(wordDoc, student.lastName, WordAPI.Extension.pdf);

            // Delete the signature, because there should be no signature in the Word file
            signature.Delete();
            SaveAs(wordDoc, student.lastName, WordAPI.Extension.docx);
        }

        //Horrible piece of shit
        private WordAPI.Picture AddSignature(WordAPI.Document wordDoc)
        {
            WordAPI.Picture signature = null;
            if (_pathToTeacherSignature != null || _pathToTeacherSignature != string.Empty)
            {
                WordAPI.Picture basePicture = wordDoc.AddPicture(mark: "[Signature]", _filePathToBaseSignature);
                //15 pt is needed for the correct location of the signature
                basePicture.Top += 15;

                signature = wordDoc.AddPicture("[NS]", _pathToTeacherSignature);
                signature.Top = basePicture.Top;
                signature.Left = basePicture.Left;
                basePicture.Delete();
            }
            else
            {
                wordDoc.Replace("[Signature]", "");
            }

            return signature;
        }

        private void SaveAs(WordAPI.Document wordDoc, string lastNameOfStudent, WordAPI.Extension extension = WordAPI.Extension.pdf)
        {
            string directory = "pdf";
            if (extension == WordAPI.Extension.docx)
            {
                directory = "word";
            }

            string fileNameWithoutExtension = "Рецензия_" + lastNameOfStudent;
            string absolutePathToFileWithoutExtension = @_folderPath + "\\" + directory + "\\" + fileNameWithoutExtension;
            wordDoc.SaveAs(absolutePathToFileWithoutExtension, extension);
        }
    }
}