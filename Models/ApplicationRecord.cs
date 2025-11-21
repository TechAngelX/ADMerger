// Models/ApplicationRecord.cs

using CsvHelper.Configuration.Attributes;

namespace ADMerger.Models
{
    public class ApplicationRecord
    {
        [Name("Applicant ID")]
        public string ApplicantID { get; set; }
       
        public string Programme { get; set; }
        public string Forename { get; set; }
        public string Surname { get; set; }
        public string Gender { get; set; }
        
        [Name("Date of Birth")]
        public string DateOfBirth { get; set; }
        
        [Name("Fee Status")]
        public string FeeStatus { get; set; }
        
        [Name("Country of Nationality")]
        public string CountryOfNationality { get; set; }
        
        [Name("Qualification name")]
        public string QualificationName { get; set; }
        
        [Name("Degree subject")]
        public string DegreeSubject { get; set; }
        
        [Name("Institution name")]
        public string InstitutionName { get; set; }
        
        [Name("Country of study")]
        public string CountryOfStudy { get; set; }
        
        [Name("Overall  grade/GPA")]
        public string OverallGradeGPA { get; set; }
       
        [Name("Equivalency note")]
        public string EquivalencyNote { get; set; }
    }
}