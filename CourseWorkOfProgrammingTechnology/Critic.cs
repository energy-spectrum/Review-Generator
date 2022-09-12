using System;
using System.Collections.Generic;

namespace CourseWorkOfProgrammingTechnology
{
    class Critic
    {
        private WordAPI.Document _wordDoc;
        private int _ratingOfStudent;
        private Criticism _advantages;
        private Criticism _disadvantages;

        private Random _rand;

        public Critic(Criticism advantages, Criticism disadvantages)
        {
            this._advantages = advantages;
            this._disadvantages = disadvantages;

            this._rand = new Random();
        }

        public void AddAdvantagesAndDisadvantages(WordAPI.Document wordDoc, Student student)
        {
            this._wordDoc = wordDoc;
            this._ratingOfStudent = student.rating;

            List<Comment> selectedAdvantages = AddAdvantages();
            HashSet<int> colorUsed = IdentifyColorsUsed(selectedAdvantages);
            AddDisadvantages(colorUsed);
        }

        private List<Comment> AddAdvantages()
        {
            List<Comment> selectedAdvantages = ChooseAdvantages();
            const string mark = "[nextAdvantage]";
            AddCriticism(selectedAdvantages, mark);

            return selectedAdvantages;
        }

        private HashSet<int> IdentifyColorsUsed(List<Comment> selectedAdvantages)
        {
            HashSet<int> colorsUsed = new HashSet<int>();
            for (int i = 0; i < selectedAdvantages.Count; i++)
            {
                if (Comment.defaultColor != selectedAdvantages[i].color)
                {
                    colorsUsed.Add(selectedAdvantages[i].color);
                }
            }

            return colorsUsed;
        }

        private void AddDisadvantages(HashSet<int> colorsUsed)
        {
            List<Comment> selectedDisadvantages = ChooseDisadvantages(colorsUsed);
            const string mark = "[nextDisadvantage]";
            AddCriticism(selectedDisadvantages, mark);
        }

        private void AddCriticism(List<Comment> selectedComments, string mark)
        {
            for (int i = 0; i < selectedComments.Count; i++)
            {
                string next = " " + mark;
                if (i == selectedComments.Count - 1)
                {
                    next = string.Empty;
                }
                _wordDoc.Replace(mark, selectedComments[i].text + next);
            }
        }

        private List<Comment> ChooseAdvantages()
        {
            List<Comment> selectedAdvantages = new List<Comment>();

            (int, int) range = _advantages.numCommentsForRatings[_ratingOfStudent];
            int numComments = _rand.Next(range.Item1, range.Item2 + 1);
            numComments = Math.Min(numComments, _advantages[_ratingOfStudent].Count);

            Shuffle(_advantages[_ratingOfStudent]);
            // Since the comments are shuffled, we choose a random comment each time
            for (int i = 0; i < numComments; i++)
            {
                selectedAdvantages.Add(_advantages[_ratingOfStudent][i]);
            }

            return selectedAdvantages;
        }

        private List<Comment> ChooseDisadvantages(HashSet<int> colorsUsed)
        {
            List<Comment> selectedDisadvantages = new List<Comment>();

            (int, int) range = _disadvantages.numCommentsForRatings[_ratingOfStudent];
            int numComments = _rand.Next(range.Item1, range.Item2 + 1);

            Shuffle(_disadvantages[_ratingOfStudent]);
            // Since the comments are shuffled, we choose a random comment each time
            for (int i = 0; i < numComments && i < _disadvantages[_ratingOfStudent].Count; i++)
            {
                if (colorsUsed.Contains(_disadvantages[_ratingOfStudent][i].color) == false)
                {
                    selectedDisadvantages.Add(_disadvantages[_ratingOfStudent][i]);
                }
                else
                {
                    numComments++;
                }
            }

            return selectedDisadvantages;
        }

        private void Shuffle(List<Comment> comments)
        {
            for (int i = comments.Count - 1; i >= 1; i--)
            {
                int j = _rand.Next(i + 1);
                var temp = comments[j];
                comments[j] = comments[i];
                comments[i] = temp;
            }
        }
    }
}
