// Services/IGradeClassificationService.cs
// © Ricki Angel 2026 | TechAngelX

namespace ADMerger.Services
{
    public interface IGradeClassificationService
    {
        string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy, string qualificationName);
        
        string ParseUKGradeText(string gradeText);
        double? ParseGradeValue(string gradeStr);
    }
}
