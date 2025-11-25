using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ADMerger.Services
{
    public class RankingService
    {
        private readonly Dictionary<string, string> theRankings = new Dictionary<string, string>();
        private readonly List<string> theInstitutionNames = new List<string>();

        public int LoadedCount => theRankings.Count;

        public void LoadTHERankings()
        {
            try
            {
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "THE Ranking 2026.xlsx");
                
                if (!File.Exists(excelPath))
                    return;
                
                using var package = new ExcelPackage(new FileInfo(excelPath));
                var worksheet = package.Workbook.Worksheets[0];
                
                if (worksheet.Dimension == null)
                    return;
                
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var rankCell = worksheet.Cells[row, 1].Value;
                    var nameCell = worksheet.Cells[row, 2].Value;
                    
                    if (rankCell != null && nameCell != null)
                    {
                        string rank = rankCell.ToString().Trim();
                        string institutionName = nameCell.ToString().Trim();
                        
                        if (!string.IsNullOrWhiteSpace(institutionName))
                        {
                            theRankings[institutionName] = rank;
                            theInstitutionNames.Add(institutionName);
                        }
                    }
                }
            }
            catch { }
        }

        public string GetTHERanking(string institutionName)
        {
            if (string.IsNullOrWhiteSpace(institutionName))
                return "NR";
            
            if (theRankings.ContainsKey(institutionName))
                return theRankings[institutionName];
            
            string bestMatch = FindBestMatch(institutionName);
            
            return bestMatch != null ? theRankings[bestMatch] : "NR";
        }

        private string FindBestMatch(string searchName)
        {
            if (string.IsNullOrWhiteSpace(searchName))
                return null;
            
            string normalizedSearch = NormalizeInstitutionName(searchName);
            var searchTerms = ExtractKeyTerms(normalizedSearch);
            
            string bestMatch = null;
            int bestScore = 0;
            
            foreach (var candidateName in theInstitutionNames)
            {
                string normalizedCandidate = NormalizeInstitutionName(candidateName);
                int score = CalculateMatchScore(normalizedSearch, normalizedCandidate, searchTerms);
                
                if (score > bestScore && score >= 60)
                {
                    bestScore = score;
                    bestMatch = candidateName;
                }
            }
            
            return bestMatch;
        }

        private string NormalizeInstitutionName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "";
            
            name = FixEncodingIssues(name);
            name = RemoveDiacritics(name);
            name = name.ToLower();
            name = name.Replace("university of", "").Replace("the ", "");
            name = Regex.Replace(name, @"[^\w\s]", " ");
            name = Regex.Replace(name, @"\s+", " ");
            
            return name.Trim();
        }

        private string FixEncodingIssues(string text)
        {
            var replacements = new Dictionary<string, string>
            {
                {"Ã ", "à"}, {"Ã¡", "á"}, {"Ã¢", "â"}, {"Ã£", "ã"}, {"Ã¤", "ä"}, {"Ã¥", "å"},
                {"Ã¨", "è"}, {"Ã©", "é"}, {"Ãª", "ê"}, {"Ã«", "ë"},
                {"Ã¬", "ì"}, {"Ã­", "í"}, {"Ã®", "î"}, {"Ã¯", "ï"},
                {"Ã²", "ò"}, {"Ã³", "ó"}, {"Ã´", "ô"}, {"Ãµ", "õ"}, {"Ã¶", "ö"},
                {"Ã¹", "ù"}, {"Ãº", "ú"}, {"Ã»", "û"}, {"Ã¼", "ü"},
                {"Ã±", "ñ"}, {"Ã§", "ç"}
            };
            
            foreach (var pair in replacements)
                text = text.Replace(pair.Key, pair.Value);
            
            return text;
        }

        private string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
            
            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                    stringBuilder.Append(c);
            }
            
            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        private List<string> ExtractKeyTerms(string normalizedName)
        {
            var commonWords = new HashSet<string> { "of", "the", "and", "in", "at", "for", "on", "a", "an" };
            
            return normalizedName
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(word => word.Length > 2 && !commonWords.Contains(word))
                .ToList();
        }

        private int CalculateMatchScore(string search, string candidate, List<string> searchTerms)
        {
            int score = 0;
            
            if (search == candidate) return 100;
            if (candidate.Contains(search)) score += 80;
            if (search.Contains(candidate)) score += 70;
            
            int matchedTerms = searchTerms.Count(term => candidate.Contains(term));
            
            if (searchTerms.Count > 0)
            {
                int termScore = (matchedTerms * 100) / searchTerms.Count;
                score = Math.Max(score, termScore);
            }
            
            if ((search.Contains("ucl") || search.Contains("university college london")) && 
                candidate.Contains("university college london"))
                return 95;
            
            if ((search.Contains("oxford") && candidate.Contains("oxford")) ||
                (search.Contains("cambridge") && candidate.Contains("cambridge")) ||
                (search.Contains("mit") && candidate.Contains("massachusetts institute")) ||
                (search.Contains("caltech") && candidate.Contains("california institute")))
                return 90;
            
            return score;
        }
    }
}
