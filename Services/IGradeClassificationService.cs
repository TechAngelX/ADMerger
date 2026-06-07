// Services/IGradeClassificationService.cs

namespace ADMerger.Services
{
    public interface IGradeClassificationService
    {
        string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy, string qualificationName);
        string GetPreferredGradeDisplay(string rawGPA, string equivalencyNote);
        string ParseUKGradeText(string gradeText);
        double? ParseGradeValue(string gradeStr);
    }
}
