namespace CourseWorkOfProgrammingTechnology
{
    class Student
    {
        readonly public string name, lastName, nameInGenitiveCase;
        readonly public string topicWork;
        readonly public int rating;

        public string ConformityRating 
        { 
            get
            {
                string[] ratings =
                {
                    "неудовлетворительно",
                    "удовлетворительно", 
                    "хорошо",
                    "отлично"
                };
                return ratings[rating - 2];
            } 
        }

        public Student(string name, string nameInGenitiveCase, string topicWork, int rating)
        {
            this.name = name;
            this.nameInGenitiveCase = nameInGenitiveCase;
            this.lastName = name.Split()[0];
            this.topicWork = topicWork;
            this.rating = rating;
        }
    }
}
