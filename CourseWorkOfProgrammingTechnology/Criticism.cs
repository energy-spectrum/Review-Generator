using System.Collections.Generic;

namespace CourseWorkOfProgrammingTechnology
{
    class Criticism
    {
        private Dictionary<int, List<Comment>> _commentsForRatings;
        readonly public Dictionary<int, (int, int)> numCommentsForRatings;
        public List<Comment> this[int rating] { get { return _commentsForRatings[rating]; } }
        public Criticism()
        {
            _commentsForRatings = new Dictionary<int, List<Comment>>()
            {
                { 3, new List<Comment>() },
                { 4, new List<Comment>() },
                { 5, new List<Comment>() }
            };
           
            numCommentsForRatings = new Dictionary<int, (int, int)>();
        }
    }
}
