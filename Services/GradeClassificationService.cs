// Services/GradeClassificationService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ADMerger.Models;

namespace ADMerger.Services
{
    public class GradeClassificationService : IGradeClassificationService
    {
        private readonly IEquivalencyService _equivalencyService;
        
        public GradeClassificationService(IEquivalencyService equivalencyService)
        {
            _equivalencyService = equivalencyService ?? throw new ArgumentNullException(nameof(equivalencyService));
        }
        
        public string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy)
        {
            if (string.IsNullOrWhiteSpace(countryOfStudy))
                return "??";
            
            if (countryOfStudy.ToLower().Contains("united kingdom"))
            {
                if (!string.IsNullOrWhiteSpace(overallGradeGPA))
                {
                    string grade = ParseUKGradeText(overallGradeGPA);
                    if (grade != "??") return grade;
                }
                
                return DetermineUKClassificationFromNote(equivalencyNote);
            }
            
            if (!string.IsNullOrWhiteSpace(equivalencyNote))
            {
                var noteLower = equivalencyNote.ToLower();
                if (noteLower.Contains("liverpool") || noteLower.Contains("reading"))
                {
                    return ExtractGradeFromLiverpoolReadingFormat(equivalencyNote);
                }
            }
            
            double? studentGrade = ParseGradeValue(overallGradeGPA);
            
            if (studentGrade == null && !string.IsNullOrWhiteSpace(equivalencyNote))
            {
                studentGrade = ExtractGradeFromNote(equivalencyNote);
            }
            
            if (studentGrade != null && !string.IsNullOrWhiteSpace(equivalencyNote))
            {
                var thresholdsFromNote = ParseThresholdsFromEquivalencyNote(equivalencyNote);
                
                if (thresholdsFromNote.Count > 0)
                {
                    return ApplyThresholds(studentGrade.Value, thresholdsFromNote);
                }
            }
            
            if (studentGrade != null)
            {
                string cleanCountry = countryOfStudy.Trim();
                var equiv = _equivalencyService.GetEquivalency(cleanCountry);
                
                if (equiv != null)
                {
                    return ApplyEquivalencyTable(studentGrade.Value, equiv);
                }
            }
            
            return DetermineUKClassificationFromNote(equivalencyNote);
        }
        
        public string ParseUKGradeText(string gradeText)
        {
            if (string.IsNullOrWhiteSpace(gradeText)) return "??";
            
            string lower = gradeText.ToLower();
            
            if (lower.Contains("first") || lower.Contains("1st") || lower == "1.0" || lower == "1")
                return "1.0";
            
            if (lower.Contains("2.1") || lower.Contains("2:1") || lower.Contains("upper second"))
                return "2.1";
            
            if (lower.Contains("2.2") || lower.Contains("2:2") || lower.Contains("lower second"))
                return "2.2";
            
            if (lower.Contains("third") || lower.Contains("3rd") || lower == "3.0" || lower == "3")
                return "3.0";
            
            return "??";
        }
        
        public double? ParseGradeValue(string gradeStr)
        {
            if (string.IsNullOrWhiteSpace(gradeStr)) return null;
            
            gradeStr = gradeStr.TrimStart('\'').Trim();
            gradeStr = gradeStr.Replace("%", "").Trim();
            
            if (gradeStr.Contains("/"))
            {
                var parts = gradeStr.Split('/');
                if (parts.Length == 2 && 
                    double.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numerator) &&
                    double.TryParse(parts[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double denominator) &&
                    denominator != 0)
                {
                    return denominator == 100.0 ? numerator : (numerator / denominator) * 100.0;
                }
            }
            
            if (double.TryParse(gradeStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            
            return null;
        }
        
        private string ExtractGradeFromLiverpoolReadingFormat(string equivalencyNote)
        {
            var percentMatches = Regex.Matches(equivalencyNote, @"(\d+(?:\.\d+)?)%");
            double bestGrade = 0;
            
            foreach (Match match in percentMatches)
            {
                if (double.TryParse(match.Groups[1].Value, out double grade))
                {
                    if (grade > bestGrade && grade <= 100 && grade >= 30)
                        bestGrade = grade;
                }
            }
            
            if (bestGrade >= 70) return "1.0";
            if (bestGrade >= 60) return "2.1";
            if (bestGrade >= 50) return "2.2";
            if (bestGrade >= 40) return "3.0";
            
            return "??";
        }
        
        private double? ExtractGradeFromNote(string equivalencyNote)
        {
            var gradeMatches = Regex.Matches(equivalencyNote, @"(?:grade of|average of|GPA of)\s*([0-9.]+(?:/[0-9.]+)?%?)");
            
            if (gradeMatches.Count > 0)
            {
                string gradeText = gradeMatches[gradeMatches.Count - 1].Groups[1].Value;
                return ParseGradeValue(gradeText);
            }
            
            return null;
        }
        
        private Dictionary<string, double> ParseThresholdsFromEquivalencyNote(string equivalencyNote)
        {
            var thresholds = new Dictionary<string, double>();
            var regex = new Regex(@"(2\.2|2\.1|1st):\s*Bachelors\s*@\s*([0-9.]+)%");
            var matches = regex.Matches(equivalencyNote);
            
            foreach (Match match in matches)
            {
                string grade = match.Groups[1].Value;
                if (double.TryParse(match.Groups[2].Value, out double threshold))
                    thresholds[grade] = threshold;
            }
            
            return thresholds;
        }
        
        private string ApplyThresholds(double studentGrade, Dictionary<string, double> thresholds)
        {
            if (thresholds.ContainsKey("1st") && studentGrade >= thresholds["1st"])
                return "1.0";
            
            if (thresholds.ContainsKey("2.1") && studentGrade >= thresholds["2.1"])
                return "2.1";
            
            if (thresholds.ContainsKey("2.2") && studentGrade >= thresholds["2.2"])
                return "2.2";
            
            return "3.0";
        }
        
        private string ApplyEquivalencyTable(double studentGrade, DegreeEquivalency equiv)
        {
            double? firstThreshold = ParseGradeValue(equiv.First);
            double? upperSecondThreshold = ParseGradeValue(equiv.SecondUpper);
            double? lowerSecondThreshold = ParseGradeValue(equiv.SecondLower);
            
            if (firstThreshold != null && studentGrade >= firstThreshold)
                return "1.0";
            
            if (upperSecondThreshold != null && studentGrade >= upperSecondThreshold)
                return "2.1";
            
            if (lowerSecondThreshold != null && studentGrade >= lowerSecondThreshold)
                return "2.2";
            
            return "3.0";
        }
        
        private string DetermineUKClassificationFromNote(string equivalencyNote)
        {
            if (string.IsNullOrWhiteSpace(equivalencyNote))
                return "??";
            
            var note = equivalencyNote.ToLower();
            var percentMatches = Regex.Matches(equivalencyNote, @"(\d+(?:\.\d+)?)%");
            double bestGrade = 0;
            
            foreach (Match match in percentMatches)
            {
                if (double.TryParse(match.Groups[1].Value, out double grade))
                {
                    if (grade > bestGrade && grade <= 100 && grade >= 30)
                        bestGrade = grade;
                }
            }
            
            if (bestGrade >= 70) return "1.0";
            if (bestGrade >= 60) return "2.1";
            if (bestGrade >= 50) return "2.2";
            if (bestGrade >= 40) return "3.0";
            
            if (note.Contains("overall classification was 1st") || note.Contains("hons 1st") || note.Contains("first class")) 
                return "1.0";
            if (note.Contains("overall classification was 2.1") || note.Contains("classification was 2.1") || note.Contains("upper second")) 
                return "2.1";
            if (note.Contains("overall classification was 2.2") || note.Contains("classification was 2.2") || note.Contains("lower second")) 
                return "2.2";
            if (note.Contains("overall classification was 3rd") || note.Contains("classification was 3rd") || note.Contains("third class")) 
                return "3.0";
            
            return "??";
        }
    }
}