namespace ADMerger.Models
{
    public class InTrayRecord
    {
        public string StudentNo { get; set; }
        public string ReceivedOn { get; set; }
    }

    public class ApplicationRecord
    {
        public string ApplicantID { get; set; }
        public string Programme { get; set; }
        public string Forename { get; set; }
        public string Surname { get; set; }
        public string Gender { get; set; }
        public string DateOfBirth { get; set; }
        public string FeeStatus { get; set; }
        public string CountryOfStudy { get; set; }
        public string CountryOfNationality { get; set; }
        public string QualificationName { get; set; }
        public string DegreeSubject { get; set; }
        public string InstitutionName { get; set; }
        public string OverallGradeGPA { get; set; }
        public string EquivalencyNote { get; set; }
    }

    public class OutputRecord
    {
        public string ReceivedDate { get; set; }
        public string DueDate { get; set; }
        public string StudentNo { get; set; }
        public string Programme { get; set; }
        public string Forename { get; set; }
        public string Surname { get; set; }
        public string Gender { get; set; }
        public string DateOfBirth { get; set; }
        public string FeeStatus { get; set; }
        public string CountryOfNationality { get; set; }
        public string QualificationName { get; set; }
        public string DegreeSubject { get; set; }
        public string InstitutionName { get; set; }
        public string THERanking { get; set; }
        public string CountryOfStudy { get; set; }
        public string OverallGradeGPA { get; set; }
        public string EquivalencyNote { get; set; }
        public string UKGrade { get; set; }
    }

    public class DegreeEquivalency
    {
        public string Country { get; set; }
        public string Third { get; set; }
        public string SecondLower { get; set; }
        public string SecondUpper { get; set; }
        public string First { get; set; }
    }
}
