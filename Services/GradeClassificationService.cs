// Services/GradeClassificationService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ADMerger.Models;

namespace ADMerger.Services
{
    public class GradeClassificationService : IGradeClassificationService
    {
        private readonly IEquivalencyService _equivalencyService;

        private readonly Regex _fractionRegex = new Regex(@"(\d+(?:\.\d+)?)\s*(?:/|out of|of)\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
        private readonly Regex _percentRegex = new Regex(@"(\d+(?:\.\d+)?)%", RegexOptions.IgnoreCase);
        private readonly Regex _customThresholdRegex = new Regex(@"(1st|2\.1|2\.2|3rd)[^@_:]*[:@]\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
        private readonly Regex _explicitGradeRegex = new Regex(@"(?:grade|gpa|average)[^0-9_]*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);

        public GradeClassificationService(IEquivalencyService equivalencyService)
        {
            _equivalencyService = equivalencyService ?? throw new ArgumentNullException(nameof(equivalencyService));
        }
        
     public string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy, string qualificationName)
     {
         // 1. Masters Check
         if (!string.IsNullOrWhiteSpace(qualificationName) && 
             qualificationName.Contains("Masters", StringComparison.OrdinalIgnoreCase))
         {
             if (!DoesNoteLookLikeUndergrad(equivalencyNote))
             {
                 return "Masters";
             }
         }
     
         // 2. Define UK-System Keywords (Universities with joint/TNE programs)
         var ukKeywords = new[] 
         { 
             "Liverpool", "Nottingham", "Exeter", "Birmingham", "Edinburgh", 
             "Reading", "Sussex", "UK Degree" 
         };
     
         // 3. Logic: Is it a UK degree regardless of the "Country of Study"?
         bool isLikelyUKDegree = IsUK(countryOfStudy) || 
                                 ukKeywords.Any(k => equivalencyNote.Contains(k, StringComparison.OrdinalIgnoreCase));
     
         if (isLikelyUKDegree)
         {
             // Try to get numeric grade from GPA field first, then the Note
             double? ukGradeValue = ExtractGradeValue(overallGradeGPA) ?? ExtractGradeValue(equivalencyNote);
             
             if (ukGradeValue.HasValue)
             {
                 double val = ukGradeValue.Value;
     
                 // Handle standard UK 100-point scale (Percentage)
                 if (val >= 35) 
                 {
                     if (val >= 70) return "1.0";
                     if (val >= 60) return "2.1";
                     if (val >= 50) return "2.2";
                     if (val >= 40) return "3.0";
                     return "Fail";
                 }
                 
                 // Handle GPA scale (e.g., 3.6/4.0) if the UK program uses it for entry
                 double normalized = GuessScaleAndNormalize(val);
                 return ApplyStandardThresholds(normalized);
             }
             
             return DetermineDomesticGrade(equivalencyNote);
         }
     
         double? studentGrade = ExtractGradeValue(equivalencyNote);
     
         if (studentGrade.HasValue)
         {
             var customThresholds = ParseCustomThresholdsFromNote(equivalencyNote);
             if (customThresholds.Count > 0)
             {
                 bool rulesAreHighScale = customThresholds.Values.Any(v => v > 20.0);
                 double gradeToTest = studentGrade.Value;
     
                 if (rulesAreHighScale)
                 {
                     gradeToTest = GuessScaleAndNormalize(studentGrade.Value);
                 }
     
                 return ApplyCustomThresholds(gradeToTest, customThresholds);
             }
     
             var equiv = _equivalencyService.GetEquivalency(countryOfStudy);
             if (equiv != null)
             {
                 return ApplySmartEquivalency(studentGrade.Value, equiv);
             }
             
             double stdNormalized = GuessScaleAndNormalize(studentGrade.Value);
             return ApplyStandardThresholds(stdNormalized);
         }
     
         return DetermineClassificationFromTextKeywords(equivalencyNote);
     }
       

        private bool DoesNoteLookLikeUndergrad(string note)
        {
            if (string.IsNullOrWhiteSpace(note)) return false;
            string lower = note.ToLowerInvariant();
            
            return lower.Contains("2:1") || lower.Contains("2.1") || 
                   lower.Contains("2:2") || lower.Contains("2.2") ||
                   lower.Contains("1st") || lower.Contains("first class") ||
                   lower.Contains("upper second") || lower.Contains("lower second") ||
                   lower.Contains("third class");
        }

        private bool IsUK(string country)
        {
            if (string.IsNullOrWhiteSpace(country)) return false;
            return country.Contains("United Kingdom", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("England", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("Scotland", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("Wales", StringComparison.OrdinalIgnoreCase);
        }

        private string DetermineDomesticGrade(string note)
        {
            double? numericGrade = ExtractGradeValue(note);
            
            if (numericGrade.HasValue)
            {
                double val = numericGrade.Value;
                if (val >= 70.0) return "1.0";
                if (val >= 60.0) return "2.1";
                if (val >= 50.0) return "2.2";
                if (val >= 40.0) return "3.0";
            }

            return ParseUKGradeText(note);
        }

        private Dictionary<string, double> ParseCustomThresholdsFromNote(string note)
        {
            var thresholds = new Dictionary<string, double>();
            if (string.IsNullOrWhiteSpace(note)) return thresholds;

            var matches = _customThresholdRegex.Matches(note);
            foreach (Match match in matches)
            {
                string label = match.Groups[1].Value.ToLower(); 
                if (double.TryParse(match.Groups[2].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                {
                    if (label.Contains("1st")) thresholds["1.0"] = value;
                    else if (label.Contains("2.1")) thresholds["2.1"] = value;
                    else if (label.Contains("2.2")) thresholds["2.2"] = value;
                    else if (label.Contains("3rd")) thresholds["3.0"] = value;
                }
            }
            return thresholds;
        }

        private string ApplyCustomThresholds(double grade, Dictionary<string, double> thresholds)
        {
            if (thresholds.ContainsKey("1.0") && grade >= thresholds["1.0"]) return "1.0";
            if (thresholds.ContainsKey("2.1") && grade >= thresholds["2.1"]) return "2.1";
            if (thresholds.ContainsKey("2.2") && grade >= thresholds["2.2"]) return "2.2";
            if (thresholds.ContainsKey("3.0") && grade >= thresholds["3.0"]) return "3.0";
            return "3.0"; 
        }

        private string ApplySmartEquivalency(double grade, DegreeEquivalency equiv)
        {
            double? t1st = ExtractGradeValue(equiv.First);       
            double? t21  = ExtractGradeValue(equiv.SecondUpper); 
            double? t22  = ExtractGradeValue(equiv.SecondLower); 

            if (!t1st.HasValue || !t22.HasValue) return "??";

            bool lowerIsBetter = t1st.Value < t22.Value;

            if (lowerIsBetter)
            {
                if (grade <= t1st.Value) return "1.0";
                if (t21.HasValue && grade <= t21.Value) return "2.1";
                if (grade <= t22.Value) return "2.2";
                return "3.0"; 
            }
            else
            {
                if (grade >= t1st.Value) return "1.0";
                if (t21.HasValue && grade >= t21.Value) return "2.1";
                if (grade >= t22.Value) return "2.2";
                return "3.0";
            }
        }

        private double? ExtractGradeValue(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            
            text = text.Trim().Replace("'", "").Replace("<", "");

            var explicitMatch = _explicitGradeRegex.Match(text);
            if (explicitMatch.Success && double.TryParse(explicitMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double extracted))
            {
                return extracted;
            }

            var fraction = _fractionRegex.Match(text);
            if (fraction.Success && 
                double.TryParse(fraction.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double num) &&
                double.TryParse(fraction.Groups[2].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double denom) &&
                denom != 0)
            {
                return num; 
            }

            var percent = _percentRegex.Match(text);
            if (percent.Success && double.TryParse(percent.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double pVal))
            {
                return pVal;
            }

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }

            return null;
        }

        private double GuessScaleAndNormalize(double grade)
        {
            if (grade > 20) return grade;
            if (grade <= 4.0) return (grade / 4.0) * 100.0;
            if (grade <= 5.0) return (grade / 5.0) * 100.0;
            if (grade <= 10.0) return (grade / 10.0) * 100.0;
            return grade * 5.0; 
        }

        private string ApplyStandardThresholds(double percent)
        {
            if (percent >= 70) return "1.0";
            if (percent >= 60) return "2.1";
            if (percent >= 50) return "2.2";
            return "3.0";
        }

        private string DetermineClassificationFromTextKeywords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "??";
            string lower = text.ToLowerInvariant();

            if (lower.Contains("@")) return "??"; 

            if (lower.Contains("first class") || lower.Contains("1st class")) return "1.0";
            if (lower.Contains("upper second") || lower.Contains("2:1") || lower.Contains("2.1")) return "2.1";
            if (lower.Contains("lower second") || lower.Contains("2:2") || lower.Contains("2.2")) return "2.2";
            if (lower.Contains("third class") || lower.Contains("3rd class")) return "3.0";

            if (lower.Contains("summa cum laude") || lower.Contains("high distinction")) return "1.0";
            if (lower.Contains("magna cum laude") || lower.Contains("distinction")) return "1.0"; 
            if (lower.Contains("cum laude") || lower.Contains("merit")) return "2.1";
            
            return "??";
        }

        public string ParseUKGradeText(string t) => DetermineClassificationFromTextKeywords(t);
        public double? ParseGradeValue(string s) => ExtractGradeValue(s);
    }
}
