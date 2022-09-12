namespace CourseWorkOfProgrammingTechnology
{
    class Student
    {
        readonly public string name, lastName;
        readonly public string topicWork;
        readonly public int rating;

        public Student(string name, string topicWork, int rating)
        {
            this.name = name;
            this.lastName = name.Split()[0];
            this.topicWork = topicWork;
            this.rating = rating;
        }

        public string GetRatingInVerbalForm()
        {
            const int minRating = 2;
            string[] ratings =
            {
                "неудовлетворительно", //2
                "удовлетворительно",
                "хорошо",
                "отлично"              //5
            };

            return ratings[rating - minRating];
        }
    }
}
