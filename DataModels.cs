// ./DataModels.cs

using CsvHelper.Configuration.Attributes;

namespace ADMerger
{
    public class InTrayRecord
    {
        [Name("Received on")]
        public string ReceivedOn { get; set; }
       
        [Name("Student No.")]
        public string StudentNo { get; set; }
       
        [Name("Name")]
        public string Name { get; set; }
    }
   
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
   
    public class OutputRecord
    {
        public string ReceivedDate { get; set; }
        public string DueDate { get; set; } // <--- FINALIZED PROPERTY
        public string StudentNo { get; set; }
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
        public string THERanking { get; set; }  // NEW PROPERTY
        public string OverallGradeGPA { get; set; }
        public string EquivalencyNote { get; set; }
        public string UKGrade { get; set; }
    }
}