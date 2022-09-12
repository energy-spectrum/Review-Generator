using Excel = Microsoft.Office.Interop.Excel;

namespace CourseWorkOfProgrammingTechnology
{
    class Comment
    {
        static public int defaultColor = (int)Excel.XlColorIndex.xlColorIndexNone;
        readonly public string text;
        readonly public int color;

        public Comment(string text, int color)
        {
            if (text[^1] != '.') { text += '.'; }
            this.text = text;
            this.color = color;
        }

        public bool AreMutuallyExclusive(Comment other)
        {
            return this.color == other.color;
        }
    }
}
