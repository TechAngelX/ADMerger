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

        // Regex helpers
        private readonly Regex _fractionRegex = new Regex(@"(\d+(?:\.\d+)?)\s*(?:/|out of|of)\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
        private readonly Regex _percentRegex = new Regex(@"(\d+(?:\.\d+)?)%", RegexOptions.IgnoreCase);
        // Pattern to find rules like "2.1: Bachelors @ 85%"
        private readonly Regex _customThresholdRegex = new Regex(@"(1st|2\.1|2\.2|3rd)[^@]*@\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);

        public GradeClassificationService(IEquivalencyService equivalencyService)
        {
            _equivalencyService = equivalencyService ?? throw new ArgumentNullException(nameof(equivalencyService));
        }
        
        public string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy, string qualificationName)
        {
            // ---------------------------------------------------------
            // 1. MASTERS CHECK (With "Integrated Masters" Sanity Check)
            // ---------------------------------------------------------
            if (!string.IsNullOrWhiteSpace(qualificationName) && 
                qualificationName.Contains("Masters", StringComparison.OrdinalIgnoreCase))
            {
                // SANITY CHECK: 
                // If the note says "2:1" or "First Class", it is likely an Integrated Masters (MEng/MSci).
                // In this case, we ABANDON the Masters logic and treat it as a normal Undergraduate degree.
                if (!DoesNoteLookLikeUndergrad(equivalencyNote))
                {
                    // It's a real Postgraduate Masters -> Just return "Masters"
                    return "Masters";
                }
            }

            // ---------------------------------------------------------
            // 2. UK UNDERGRADUATE LOGIC
            // ---------------------------------------------------------
            if (IsUK(countryOfStudy))
            {
                return DetermineDomesticGrade(overallGradeGPA, equivalencyNote);
            }

            // ---------------------------------------------------------
            // 3. INTERNATIONAL LOGIC
            // ---------------------------------------------------------
            
            // Step A: Get a raw number (from Grade column or Notes)
            double? studentGrade = ExtractGradeValue(overallGradeGPA) ?? ExtractGradeValue(equivalencyNote);

            if (studentGrade.HasValue)
            {
                // Step B: Check for explicit rules in the text (e.g. "2.1 @ 85%")
                var customThresholds = ParseCustomThresholdsFromNote(equivalencyNote);
                if (customThresholds.Count > 0)
                {
                    double normalized = GuessScaleAndNormalize(studentGrade.Value);
                    return ApplyCustomThresholds(normalized, customThresholds);
                }

                // Step C: Check the Country Database (e.g. Austria logic)
                var equiv = _equivalencyService.GetEquivalency(countryOfStudy);
                if (equiv != null)
                {
                    return ApplySmartEquivalency(studentGrade.Value, equiv);
                }
                
                // Step D: Fallback (Standard 0-100 scale)
                double stdNormalized = GuessScaleAndNormalize(studentGrade.Value);
                return ApplyStandardThresholds(stdNormalized);
            }

            // Step E: Last Resort - Look for keywords ("Distinction", "Merit")
            return DetermineClassificationFromTextKeywords(equivalencyNote);
        }

        // --- HELPER LOGIC ---

        private bool DoesNoteLookLikeUndergrad(string note)
        {
            if (string.IsNullOrWhiteSpace(note)) return false;
            string lower = note.ToLowerInvariant();
            
            // If these terms appear, use 1st/2:1 logic, NOT the "Masters" label
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

        private string DetermineDomesticGrade(string grade, string note)
        {
            // 1. Check Numbers (Most accurate for UK)
            double? numericGrade = ExtractGradeValue(grade) ?? ExtractGradeValue(note);
            
            if (numericGrade.HasValue)
            {
                double val = numericGrade.Value;
                if (val >= 70.0) return "1.0";
                if (val >= 60.0) return "2.1";
                if (val >= 50.0) return "2.2";
                if (val >= 40.0) return "3.0";
            }

            // 2. Check Text Keywords
            string textResult = ParseUKGradeText(grade);
            if (textResult != "??") return textResult;

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
            // Custom rules usually imply "Higher is Better"
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

            // CHECK DIRECTION: Is this a "Reverse Scale" country? (e.g. Austria)
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
            
            // Clean specific garbage from your file (<, ', etc)
            text = text.Trim().Replace("'", "").Replace("<", "");

            // 1. Try "X / Y"
            var fraction = _fractionRegex.Match(text);
            if (fraction.Success && 
                double.TryParse(fraction.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double num) &&
                double.TryParse(fraction.Groups[2].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double denom) &&
                denom != 0)
            {
                return num; 
            }

            // 2. Try "%"
            var percent = _percentRegex.Match(text);
            if (percent.Success && double.TryParse(percent.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double pVal))
            {
                return pVal;
            }

            // 3. Try "current grade of 90.26"
            var currentGradeMatch = Regex.Match(text, @"(?:grade|gpa|average)[^0-9]*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
            if (currentGradeMatch.Success && double.TryParse(currentGradeMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double extracted))
            {
                return extracted;
            }

            // 4. Try raw number
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            {
                return val;
            }

            return null;
        }

        private double GuessScaleAndNormalize(double grade)
        {
            if (grade > 20) return grade; // Already %
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

            // Ignore "2.1" if it's part of a rule string like "2.1 @ 85%"
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

        // Interface compatibility
        public string ParseUKGradeText(string t) => DetermineClassificationFromTextKeywords(t);
        public double? ParseGradeValue(string s) => ExtractGradeValue(s);
    }
}
