using System.Collections.Generic;

namespace CourseWorkOfProgrammingTechnology
{
    struct InitialData
    {
        public GeneralData generalData;
        public Criticism advantages;
        public Criticism disadvantages;
        public List<Student> students;

        public InitialData(GeneralData generalData, Criticism advantages, Criticism disadvantages, List<Student> students)
        {
            this.generalData = generalData;
            this.advantages = advantages;
            this.disadvantages = disadvantages;
            this.students = students;
        }
    }
}
