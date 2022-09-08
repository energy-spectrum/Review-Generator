namespace CourseWorkOfProgrammingTechnology
{
    class Date
    {
        readonly public string day, month, year;
        public Date(string day, string month, string year)
        {
            this.day = day;
            this.month = month;
            this.year = year;
        }
        public Date(string sDate)
        {
            //string[] months =
            //{
            //    "январь",
            //    "февраль",
            //    "март",
            //    "апрель",
            //    "май",
            //    "июнь",
            //    "июль",
            //    "август",
            //    "сеньтябрь",
            //    "октябрь",
            //    "ноябрь",
            //    "декабрь"
            //};
            string[] aDate = sDate.Split('.');
            this.day = aDate[0];
            if (this.day.Length == 1) { this.day = "0" + this.day; }

            this.month = aDate[1];
            if (this.month.Length == 1) { this.month = "0" + this.month; }

            //Это значит, что в дату залезло и время
            if (aDate[2].Contains(' '))
                this.year = aDate[2].Split(' ')[0];
            else
                this.year = aDate[2];
        }
    }

    class GeneralData
    {
        private string _faculty,
                      _type,
                      _course,
                      _directionOfTraining,
                      _directivity,
                      _teacher,
                      _academicTitleAndPosition,
                      _group,
                      _pathToTeacherSignature;

        private Date _date;

        private string _lastNameTeacher,
                      _firstNameTeacher,
                      _patronymicTeacher,
                      _initialsAndSurnameTeacher;

        public string Faculty
        {
            get { return _faculty; }
            set { if (_faculty == null) { _faculty = value; } }
        }
        public string Type 
        {
            get { return _type; }
            set { if (_type == null) { _type = value; } }
        }
        public string Course
        {
            get { return _course; }
            set{ if (_course == null) {  _course = value; } }
        }
        public string DirectionOfTraining
        {
            get { return _directionOfTraining; }
            set { if (_directionOfTraining == null) { _directionOfTraining = value; } }
        }
        public string Directivity
        {
            get { return _directivity; }
            set { if (_directivity == null) {_directivity = value; } }
        }
        public string AcademicTitleAndPosition
        {
            get { return _academicTitleAndPosition; }
            set { if (_academicTitleAndPosition == null) { _academicTitleAndPosition = value; } }
        }
        public string Group
        {
            get { return _group; }
            set { if (_group == null) { _group = value;} }
        }

        public Date Date
        {
            get { return _date; }
            set { if (_date == null) { _date = value; } }
        }
       
        public string PathToTeacherSignature
        {
            get { return _pathToTeacherSignature; }
            set { if (_pathToTeacherSignature == null) { _pathToTeacherSignature = value; } }
        }

        public string Teacher
        {
            get { return _teacher; }
            set
            {
                if(_teacher != null) { return; }

                _teacher = value;
                string[] name = _teacher.Split();
                LastNameTeacher = name[0];
                FirstNameTeacher = name[1];
                PatronymicTeacher = name[2];

                InitialsAndSurnameTeacher = _firstNameTeacher.Substring(0, 1) + "." +
                                            _patronymicTeacher.Substring(0, 1) + "." + " " +
                                            _lastNameTeacher;
            }
        }

        public string LastNameTeacher
        {
            get { return _lastNameTeacher; }
            private set { if (_lastNameTeacher == null) { _lastNameTeacher = value; } }
        }
        public string FirstNameTeacher
        {
            get { return _firstNameTeacher; }
            private set { if (_firstNameTeacher == null) { _firstNameTeacher = value; } }
        }
        public string PatronymicTeacher
        {
            get { return _patronymicTeacher; }
            private set { if (_patronymicTeacher == null) { _patronymicTeacher = value; } }
        }
        public string InitialsAndSurnameTeacher
        {
            get { return _initialsAndSurnameTeacher; }
            private set { if (_initialsAndSurnameTeacher == null) { _initialsAndSurnameTeacher = value; } }
        }
    }
}
