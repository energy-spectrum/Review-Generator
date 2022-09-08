using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace CourseWorkOfProgrammingTechnology
{
    class GeneratingReviews
    {
        private GeneralData _generalData;
        private Criticism _advantages;
        private Criticism _disadvantages;
        private List<Student> _students;
        private Random _rand;

        private Object oTemplate;                                                                                                                       // = Projectname.Properties.Resourses.Rewio.doc;//@"Resources\Rewio.doc";// D:\Programming\C#_Projects\ProgrammingTechnologies\CourseWorkOfProgrammingTechnology\CourseWorkOfProgrammingTechnology\Resources\Rewio.doc";
        private string _filePathBaseSignature;

        public GeneratingReviews(GeneralData generalData, Criticism advantages, Criticism disadvantages, List<Student> students)
        {
            this._generalData = generalData;
            this._advantages = advantages;
            this._disadvantages = disadvantages;
            this._students = students;
            this._rand = new Random();

            string assmblypath = Assembly.GetEntryAssembly().Location;
            string appPath = Path.GetDirectoryName(assmblypath);

            this.oTemplate = appPath + @"\Properties\Rewio.doc";
            this._filePathBaseSignature = appPath + @"\Properties\BaseSignature.jpg";
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

        private Object oMissing = System.Reflection.Missing.Value;
        private void SaveDocumentIntoPdf(Word.Document wordDoc, string fileNameWithoutExtension)
        {
            object outputFileName = fileNameWithoutExtension + ".pdf";
            object fileFormat = Word.WdSaveFormat.wdFormatPDF;

            // Save document into PDF Format
            wordDoc.SaveAs(ref outputFileName,
                        ref fileFormat, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }

        private void CreateReview(Word.Application wordApp, Student student, string folderPath)
        {
            Word.Document wordDoc = null;
            try
            {
                // Add a new document
                wordDoc = new Word.Document();
                wordDoc = wordApp.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);


                Replace(wordApp, "faculty", _generalData.Faculty);
                if(_generalData.Type == "Работа")
                {
                    Replace(wordApp, "курсовой проект", "курсовую работу");
                    Replace(wordApp, "курсового проекта", "курсовой работы");
                }
                Replace(wordApp, "Brattsev", student.name);
                Replace(wordApp, "__X__", "__" + _generalData.Course + "__");
                Replace(wordApp, "directionOfTraining", _generalData.DirectionOfTraining);
                Replace(wordApp, "directivity", _generalData.Directivity);
                Replace(wordApp, "topicWork", student.topicWork);
                Replace(wordApp, "academicTitleAndPosition", _generalData.AcademicTitleAndPosition);
                Replace(wordApp, "ФИО_Преподавателя", _generalData.Teacher);
                Replace(wordApp, "rating", student.ConformityRating);

                List<Comment> lAdvantages = ChooseAdvantages(student.rating);
                for (int i = 0; i < lAdvantages.Count; i++)
                {
                    string next = " " + "nextAdvantage";
                    if (i == lAdvantages.Count - 1) { next = string.Empty; }
                    Replace(wordApp, "nextAdvantage", lAdvantages[i].text + next);
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
                    Replace(wordApp, "nextDisadvantage", lDisadvantages[i].text + next);
                }
                Replace(wordApp, "5", student.rating.ToString());
                Replace(wordApp, "И.О. ФамилияРецензента", _generalData.InitialsAndSurnameTeacher);
                Replace(wordApp, "day", _generalData.Date.day);
                Replace(wordApp, "month", _generalData.Date.month);
                Replace(wordApp, "year", _generalData.Date.year);

                Word.Shape signature = null;
                if (_generalData.PathToTeacherSignature != null || _generalData.PathToTeacherSignature != string.Empty)
                {
                    signature = AddSignature(wordApp, _generalData.PathToTeacherSignature);
                }
                else
                {
                    Replace(wordApp, "Signature", "");
                }

                string fileNameWithoutExtension = @folderPath + "\\pdf" + "\\Рецензия_" + student.lastName;
                SaveDocumentIntoPdf(wordDoc, fileNameWithoutExtension);

                // Delete the signature, because there should be no signature in the Word file
                DelateSignature(signature);
                fileNameWithoutExtension = @folderPath + "\\word" + "\\Рецензия_" + student.lastName;
                // Save the document
                Object oSaveAsFile = (Object)(fileNameWithoutExtension + ".docx");
                wordDoc.SaveAs
                (
                    ref oSaveAsFile, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing
                );
            }
            finally
            {
                Object oSaveChanges = false;
                // Close the file
                wordDoc.Close(ref oSaveChanges, ref oMissing, ref oMissing);
                wordDoc = null;
            }
        }
        
        public void CreateReviews(string folderPath)
        {
            Directory.CreateDirectory(folderPath + "\\pdf");
            Directory.CreateDirectory(folderPath + "\\word");

            Word.Application wordApp = null;
            try
            {
                // Create object of Word
                wordApp = new Word.Application();
                // Don't display Word
                wordApp.Visible = false;

                for (int idxStudent = 0; idxStudent < _students.Count; idxStudent++)
                {
                    CreateReview(wordApp, _students[idxStudent], folderPath);
                }
            }
            finally
            {
                // Quit Word.exe
                wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                wordApp = null;
            }
        }

        private void Replace(Word.Application wordApp, string oldText, string newText, Word.WdReplace replace = Word.WdReplace.wdReplaceAll)
        {
            // Clearing the search parameters
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Replacement.ClearFormatting();

            Object findText = oldText;//<Replacing which text>
            Object replaceWith = newText;//<Replacing with what text>
            Object oReplace = replace;//<How many such occurrences are we changing / replace = Word.WdReplace.wdReplaceAll>

            Object oMissing = System.Reflection.Missing.Value;

            wordApp.Selection.Find.Execute(ref findText, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref replaceWith,
            ref oReplace, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }

        private Word.Shape AddSignature(Word.Application wordApp, string pathToTeacherSignature)
        {
            var sel = wordApp.Selection;
            sel.Find.ClearFormatting();
            sel.Find.Replacement.ClearFormatting();

            sel.Find.Text = "Signature";
            sel.Find.Execute(Replace: Word.WdReplace.wdReplaceNone);
            sel.Range.Select();

            //This code inserts the imageSignature
            Word.InlineShape signatureInlineShape = sel.InlineShapes.AddPicture(
                FileName: _filePathBaseSignature,//"C:\\Users\\batar\\Desktop\\ПодписьРецензента.jpg",
                LinkToFile: false,
                SaveWithDocument: true);
            var baseSignatureShape = signatureInlineShape.ConvertToShape();
            baseSignatureShape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;
            baseSignatureShape.Top += 15f;

            ///////////////////////////////////////////////////
            sel.Find.ClearFormatting();
            sel.Find.Replacement.ClearFormatting();

            sel.Find.Text = "NS";
            sel.Find.Execute(Replace: Word.WdReplace.wdReplaceNone);
            sel.Range.Select();

            //This code inserts the imageSignature
            Word.InlineShape signatureInlineShape1 = sel.InlineShapes.AddPicture(
                FileName: pathToTeacherSignature,
                LinkToFile: false,
                SaveWithDocument: true);
            var signatureShape = signatureInlineShape1.ConvertToShape();
            signatureShape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;

            signatureShape.Top = baseSignatureShape.Top;
            signatureShape.Left = baseSignatureShape.Left;
           
            DelateSignature(baseSignatureShape);

            return signatureShape;
        }
        private void DelateSignature(Word.Shape signature)
        {
            signature?.Delete();
        }
    }
}


//const float c = 0.0352778f;

//signatureShape.Height = baseSignatureShape.Height;
//signatureShape.Width = baseSignatureShape.Width;//3.45f * 29f;

// signatureShape.Width = baseSignatureShape.Width;
// signatureShape.Height = baseSignatureShape.Height;
