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
        private readonly Regex _priorityKeywordRegex = new Regex(@"(?:average|final|cgpa|overall|gpa)\s*[:\-\s]*\s*(\d{1,3}(?:\.\d+)?)", RegexOptions.IgnoreCase);
        private readonly Regex _generalGradeRegex = new Regex(@"\b(?!(?:1\.1|2\.1|2\.2))\d{1,3}(?:\.\d+)?\b");

        public GradeClassificationService(IEquivalencyService equivalencyService)
        {
            _equivalencyService = equivalencyService ?? throw new ArgumentNullException(nameof(equivalencyService));
        }

        public string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy, string qualificationName)
        {
            if (string.IsNullOrWhiteSpace(equivalencyNote)) return "??";

            // SCRUB TREND: Remove underscores and double underscores used as delimiters in your data
            string cleanNote = equivalencyNote.Replace("__", " ").Replace("_", " ");

            // 0. Masters Shield
            if (qualificationName?.Contains("Masters", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (!DoesNoteLookLikeUndergrad(cleanNote)) return "Masters";
            }

            // 1. KEYWORD PRIORITY (Trust "1st", "2.1" in text before math)
            string keywordResult = DetermineClassificationFromTextKeywords(cleanNote);
            if (keywordResult != "??") return keywordResult;

            double? studentGrade = ExtractGradeValue(cleanNote);
            string noteLower = cleanNote.ToLowerInvariant();

            // 2. HARD EXCEPTIONS (Glasgow, Lancaster, Italy)
            if (noteLower.Contains("lancaster") || noteLower.Contains("glasgow"))
            {
                double val = studentGrade ?? 0;
                if (val >= 17.5) return "1.0";
                if (val >= 14.5) return "2.1";
                if (val >= 11.5) return "2.2";
                if (val >= 8.5)  return "3.0";
                return "Fail";
            }

            if (countryOfStudy?.Contains("Italy", StringComparison.OrdinalIgnoreCase) == true || noteLower.Contains("110"))
            {
                var m110 = Regex.Matches(cleanNote, @"\b(10[0-9]|110)\b");
                if (m110.Count > 0)
                {
                    double maxVal = m110.Cast<Match>().Select(m => double.Parse(m.Value)).Max();
                    if (maxVal >= 108) return "1.0";
                    if (maxVal >= 106) return "2.1";
                    if (maxVal >= 101) return "2.2";
                }
                if (studentGrade.HasValue && studentGrade.Value <= 30)
                {
                    double v = studentGrade.Value;
                    if (v >= 28.5) return "1.0";
                    if (v >= 26.5) return "2.1";
                    if (v >= 24.0) return "2.2";
                    return "3.0";
                }
            }

            // 3. GREEDY UK PATTERN RECOGNITION (Added Reading, Essex, Nottingham, Liverpool)
            var ukKeywords = new[] { 
                "university of", "college london", "hons", "uk degree", 
                "reading", "essex", "nottingham", "liverpool", "exeter", 
                "birmingham", "edinburgh", "sussex", "warwick", "manchester",
                "kcl", "ucl", "lse", "imperial", "oxford", "cambridge", "stirling"
            };

            bool looksLikeUK = IsUK(countryOfStudy) || ukKeywords.Any(k => noteLower.Contains(k));

            if (studentGrade.HasValue)
            {
                double val = studentGrade.Value;
                
                // UK Standard 100-point scale
                if (looksLikeUK && val >= 35) 
                {
                    if (val >= 70) return "1.0";
                    if (val >= 60) return "2.1";
                    if (val >= 50) return "2.2";
                    if (val >= 40) return "3.0";
                    return "Fail";
                }

                // International Fallback (CSV Lookup)
                var equiv = _equivalencyService.GetEquivalency(countryOfStudy);
                if (equiv != null) return ApplySmartEquivalency(val, equiv);
                
                return ApplyStandardThresholds(GuessScaleAndNormalize(val));
            }

            return "??";
        }

        private double? ExtractGradeValue(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            text = text.Trim().Replace("'", "").Replace("<", "").Replace(">", "");

            var priorityMatch = _priorityKeywordRegex.Match(text);
            if (priorityMatch.Success && double.TryParse(priorityMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double pExtracted))
                return pExtracted;

            var fraction = _fractionRegex.Match(text);
            if (fraction.Success && double.TryParse(fraction.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                return num; 

            var percent = _percentRegex.Match(text);
            if (percent.Success && double.TryParse(percent.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double pVal))
                return pVal;

            var generalMatches = _generalGradeRegex.Matches(text);
            if (generalMatches.Count > 0)
            {
                return generalMatches.Cast<Match>()
                    .Select(m => double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0)
                    .Where(d => d < 500) // SHIELD: Ignore years
                    .DefaultIfEmpty(0)
                    .Max();
            }
            return null;
        }

        private bool DoesNoteLookLikeUndergrad(string note)
        {
            if (string.IsNullOrWhiteSpace(note)) return false;
            string lower = note.ToLowerInvariant();
            return lower.Contains("2:1") || lower.Contains("2.1") || lower.Contains("2:2") || lower.Contains("2.2") || lower.Contains("1st") || lower.Contains("first class");
        }

        private bool IsUK(string country)
        {
            if (string.IsNullOrWhiteSpace(country)) return false;
            return country.Contains("United Kingdom", StringComparison.OrdinalIgnoreCase) || country.Contains("England", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("Scotland", StringComparison.OrdinalIgnoreCase) || country.Contains("Wales", StringComparison.OrdinalIgnoreCase);
        }

        private string DetermineClassificationFromTextKeywords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "??";
            // REMOVE REQUIRES 2_1: Remove the course requirements so we don't grade the student against the job description
            string lower = text.ToLowerInvariant();
            lower = lower.Replace("requires 2.1", "")
                             .Replace("requires a 2.1", "")
                             .Replace("requirement is 2.1", "")
                             .Replace("requires 1st", "");
            // Now look for the ACTUAL grade the student has
            if (Regex.IsMatch(lower, @"\b(1st|1\.0|first class)\b")) return "1.0";
            if (Regex.IsMatch(lower, @"\b(2\.1|2:1|upper second)\b")) return "2.1";
            if (Regex.IsMatch(lower, @"\b(2\.2|2:2|lower second)\b")) return "2.2";
            if (Regex.IsMatch(lower, @"\b(3\.0|3rd|third class)\b")) return "3.0";

            if (lower.Contains("summa cum laude") || lower.Contains("high distinction")) return "1.0";
            if (lower.Contains("magna cum laude") || lower.Contains("distinction")) return "1.0"; 
            if (lower.Contains("cum laude") || lower.Contains("merit")) return "2.1";
            
            return "??";
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
            if (t1st.Value < t22.Value)
            {
                if (grade <= t1st.Value) return "1.0";
                if (t21.HasValue && grade <= t21.Value) return "2.1";
                return grade <= t22.Value ? "2.2" : "3.0";
            }
            else
            {
                if (grade >= t1st.Value) return "1.0";
                if (t21.HasValue && grade >= t21.Value) return "2.1";
                return grade >= t22.Value ? "2.2" : "3.0";
            }
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
            return percent >= 50 ? "2.2" : "3.0";
        }

        public string ParseUKGradeText(string t) => DetermineClassificationFromTextKeywords(t);
        public double? ParseGradeValue(string s) => ExtractGradeValue(s);
    }
}
