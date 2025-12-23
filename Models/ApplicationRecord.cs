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
        
        [Name("Fee Status")]
        public string FeeStatus { get; set; }
        
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
        
        [Name("Grade Achieved/Pending")]
        public string GradeAchievedPending { get; set; }
    }
}
