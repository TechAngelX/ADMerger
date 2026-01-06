// Models/ApplicationRecord.cs
// © Ricki Angel 2026 | TechAngelX

using CsvHelper.Configuration.Attributes;

namespace ADMerger.Models
{
    public class ApplicationRecord
    {
        [Name("Applicant ID")]
        public string? ApplicantID { get; set; }
       
        public string? Programme { get; set; }
        public string? Forename { get; set; }
        public string? Surname { get; set; }
        
        [Name("Fee Status")]
        public string? FeeStatus { get; set; }
        
        [Name("Qualification name")]
        public string? QualificationName { get; set; }
        
        [Name("Degree subject")]
        public string? DegreeSubject { get; set; }
        
        [Name("Institution name")]
        public string? InstitutionName { get; set; }
        
        [Name("Country of study")]
        public string? CountryOfStudy { get; set; }
        
        [Name("Overall  grade/GPA")]
        public string? OverallGradeGPA { get; set; }
       
        [Name("Equivalency note")]
        public string? EquivalencyNote { get; set; }
        
        [Name("Grade Achieved/Pending")]
        public string? GradeAchievedPending { get; set; }

        // Additional fields available for optional inclusion in output
        [Name("Known as")]
        public string? KnownAs { get; set; }

        [Name("Mode of attendance")]
        public string? ModeOfAttendance { get; set; }

        public string? Location { get; set; }
        public string? State { get; set; }

        [Name("Email address")]
        public string? EmailAddress { get; set; }

        public string? Gender { get; set; }

        [Name("Date of Birth")]
        public string? DateOfBirth { get; set; }

        [Name("Country of Nationality")]
        public string? CountryOfNationality { get; set; }

        [Name("Qualification end date")]
        public string? QualificationEndDate { get; set; }

        [Name("Total Mark equivalency")]
        public string? TotalMarkEquivalency { get; set; }

        [Name("Admissions referral note")]
        public string? AdmissionsReferralNote { get; set; }

        [Name("Admissions referred to department date")]
        public string? AdmissionsReferredToDepartmentDate { get; set; }

        [Name("Department recommended decision")]
        public string? DepartmentRecommendedDecision { get; set; }

        [Name("Department recommended decision date")]
        public string? DepartmentRecommendedDecisionDate { get; set; }

        [Name("Deposit due date")]
        public string? DepositDueDate { get; set; }

        [Name("Deposit payment status")]
        public string? DepositPaymentStatus { get; set; }

        [Name("Initial decision")]
        public string? InitialDecision { get; set; }

        [Name("Initial Decision date")]
        public string? InitialDecisionDate { get; set; }

        [Name("Decision/Response")]
        public string? DecisionResponse { get; set; }

        [Name("Reply by date")]
        public string? ReplyByDate { get; set; }

        [Name("Academic year")]
        public string? AcademicYear { get; set; }

        public string? Tag { get; set; }

        [Name("ELP type")]
        public string? ELPType { get; set; }

        [Name("ELP verification status")]
        public string? ELPVerificationStatus { get; set; }
    }
}
