using System;

using Word = Microsoft.Office.Interop.Word;

namespace WordAPI
{
    class Application
    {
        private Word.Application _wordApp;

        public event Action OnQuit;

        public Application()
        {
            RunWordApplication();
        }

        private void RunWordApplication()
        {
            try
            {
                TryRunWordApplication();
            }
            catch
            {
                Quit();
                throw new Exception("Error run Word application");
            }
        }

        private void TryRunWordApplication()
        {
            _wordApp = new Word.Application();
            _wordApp.Visible = false;
        }

        public void Quit(bool saveChanges = false)
        {
            OnQuit?.Invoke();
            _wordApp?.Quit(saveChanges);
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
        private Word.Document _wordDoc;
        private Word.Selection _selection;

        public Document(Word.Application wordApp, string filePath)
        {
            this._wordDoc = wordApp.Documents.Add(Template: filePath);
            this._selection = wordApp.Selection;
        }

        public void SaveAs(string absolutePathToFileWithoutExtension, Extension extension = Extension.docx)
        {
            Word.WdSaveFormat fileFormat = Word.WdSaveFormat.wdFormatDocumentDefault;

            if (extension == Extension.pdf)
            {
                fileFormat = Word.WdSaveFormat.wdFormatPDF;
            }
            _wordDoc.SaveAs(FileName: absolutePathToFileWithoutExtension, FileFormat: fileFormat);
        }

        public void Close(bool saveChanges = false)
        {
            _wordDoc.Close(SaveChanges: saveChanges); 
            _wordDoc = null;
        }

        public void Replace(string oldText, string newText)
        {
            ClearSearchParameters();

            _selection.Find.Execute(FindText: oldText,
                                    ReplaceWith: newText,
                                    Replace: Word.WdReplace.wdReplaceAll);
        }

        public Picture AddPicture(string mark, string filePathToPicture)
        {
            ClearSearchParameters();

            _selection.Find.Text = mark;
            _selection.Find.Execute(Replace: Word.WdReplace.wdReplaceNone);
            _selection.Range.Select();

            //This code inserts the picture
            Word.InlineShape pictureInlineShape = _selection.InlineShapes.AddPicture(FileName: filePathToPicture,
                                                                                     LinkToFile: false,
                                                                                     SaveWithDocument: true);

            var pictureShape = pictureInlineShape.ConvertToShape();
            pictureShape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;

            Picture picture = new Picture(pictureShape);

            return picture;
        }

        private void ClearSearchParameters()
        {
            _selection.Find.ClearFormatting();
            _selection.Find.Replacement.ClearFormatting();
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

        public void Delete()
        {
            shape?.Delete();
            shape = null;
        }
    }
}
