using System;
using System.Diagnostics;

using Word = Microsoft.Office.Interop.Word;

namespace WordAPI
{
    class Application
    {
        private Word.Application _wordApp = null;

        private Object oMissing = System.Reflection.Missing.Value;

        public event Action OnQuit;

        public Application()
        {
            TryOpenWord();
        }

        private void TryOpenWord()
        {
            try
            {
                OpenWord();
            }
            catch
            {
                Debug.WriteLine("Error open file");
                Quit();
            }
        }

        private void OpenWord()
        {
            // Create object of Word
            _wordApp = new Word.Application();
            // Don't display Word
            _wordApp.Visible = false;
        }

        public void Quit()
        {
            OnQuit?.Invoke();
            _wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
            _wordApp = null;
        }

        public Document AddDocument(string filePath)
        {
            Document wordDoc = new Document(_wordApp, filePath);

            return wordDoc;
        }
    }

    class Document
    {
        private Word.Application _wordApp;
        private Word.Document _wordDoc;

        private Object oMissing = System.Reflection.Missing.Value;

        public Document(Word.Application wordApp, string filePath)
        {
            this._wordApp = wordApp;

            //Object oTemplate = filePath;
            _wordDoc = _wordApp.Documents.Add(Template: filePath);// ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
        }

        public Picture AddPicture(string mark, string filePathToPicture)
        {
            var sel = _wordApp.Selection;

            sel.Find.ClearFormatting();
            sel.Find.Replacement.ClearFormatting();

            sel.Find.Text = mark;
            sel.Find.Execute(Replace: Word.WdReplace.wdReplaceNone);
            sel.Range.Select();

            //This code inserts the picture
            Word.InlineShape pictureInlineShape = sel.InlineShapes.AddPicture(
                FileName: filePathToPicture,
                LinkToFile: false,
                SaveWithDocument: true);


            var pictureShape = pictureInlineShape.ConvertToShape();
            pictureShape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;

            Picture picture = new Picture(pictureShape);

            return picture;
        }

        public void Replace(string oldText, string newText)
        {
            // Clearing the search parameters
            _wordApp.Selection.Find.ClearFormatting();
            _wordApp.Selection.Find.Replacement.ClearFormatting();

            Object oFindText = oldText;//<Replacing which text>
            Object oReplaceWith = newText;//<Replacing with what text>

            Word.WdReplace replace = Word.WdReplace.wdReplaceAll;
            Object oReplace = replace;//<How many such occurrences are we changing / replace = Word.WdReplace.wdReplaceAll>

            _wordApp.Selection.Find.Execute(ref oFindText, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oReplaceWith,
            ref oReplace, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }

        public void SaveAs(string fileNameWithoutExtension, string extension = ".docx")
        {
            Object oSaveAsFile = (Object)(fileNameWithoutExtension + extension);
            Object oFileFormat = oMissing;
            if (extension == ".pdf")
            {
                oFileFormat = Word.WdSaveFormat.wdFormatPDF;
            }
            _wordDoc.SaveAs
            (
                ref oSaveAsFile, ref oFileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing
            );
        }

        public void Close(bool saveChanges = false)
        {
            Object oSaveChanges = saveChanges;
            // Close the file
            _wordDoc.Close(ref oSaveChanges, ref oMissing, ref oMissing);
            _wordDoc = null;
        }
    }

    class Picture
    {
        private Word.Shape shape;

        public float Top
        {
            get { return shape.Top; }
            set { shape.Top = value; }
        }
        public float Left
        {
            get { return shape.Left; }
            set { shape.Left = value; }
        }

        public Picture(Word.Shape shape)
        {
            this.shape = shape;
        }

        public void Delate()
        {
            shape?.Delete();
            shape = null;
        }
    }
}
