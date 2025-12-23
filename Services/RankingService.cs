// Services/RankingService.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace ADMerger.Services
{
    public class RankingService : IRankingService
    {
        private readonly Dictionary<string, string> _rankings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _institutionNames = new List<string>();
        private readonly Dictionary<string, string> _institutionMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly IInstitutionMatchingService _matchingService;
        
        public int Count => _rankings.Count;
        
        public RankingService(IInstitutionMatchingService matchingService)
        {
            _matchingService = matchingService ?? throw new ArgumentNullException(nameof(matchingService));
        }
        
        public void LoadRankings()
        {
            _rankings.Clear();
            _institutionNames.Clear();
            LoadInstitutionMappings();
            
            try
            {
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "THE Ranking 2026.xlsx");
                
                if (!File.Exists(excelPath))
                {
                    excelPath = Path.Combine(Directory.GetCurrentDirectory(), "data", "THE Ranking 2026.xlsx");
                }

                if (File.Exists(excelPath))
                {
                    using (var stream = File.OpenRead(excelPath))
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; 
                        
                        if (worksheet.Dimension != null)
                        {
                            int totalRows = worksheet.Dimension.Rows;
                            for (int row = 2; row <= totalRows; row++)
                            {
                                var rankCell = worksheet.Cells[row, 1].Value;
                                var nameCell = worksheet.Cells[row, 2].Value;
                                
                                if (rankCell != null && nameCell != null)
                                {
                                    string rank = rankCell.ToString().Trim();
                                    string name = nameCell.ToString().Trim();
                                    
                                    if (!string.IsNullOrWhiteSpace(name))
                                    {
                                        _rankings[name] = rank;
                                        _institutionNames.Add(name);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ranking Load Error: {ex.Message}");
            }
        }
        
        public string GetRanking(string institutionName)
        {
            if (string.IsNullOrWhiteSpace(institutionName)) return "NR";
            
            institutionName = NormalizeAbbreviations(institutionName);
            
            if (_rankings.ContainsKey(institutionName))
            {
                return CleanRankingText(_rankings[institutionName]);
            }
            
            string bestMatch = _matchingService.FindBestMatch(institutionName, _institutionNames);
            if (bestMatch != null)
            {
                return CleanRankingText(_rankings[bestMatch]);
            }
            
            return "NR";
        }
        
        private string CleanRankingText(string ranking)
        {
            if (string.IsNullOrWhiteSpace(ranking)) return "NR";
            return ranking.Replace(char.ConvertFromUtf32(0x2013), "-").Replace(char.ConvertFromUtf32(0x2014), "-");
        }
        
        private void LoadInstitutionMappings()
        {
            try
            {
                string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "institution_mappings.csv");
                if (!File.Exists(csvPath)) return;

                var lines = File.ReadAllLines(csvPath);
                foreach (var line in lines.Skip(1))
                {
                    var parts = line.Split(',');
                    if (parts.Length >= 2) _institutionMappings[parts[0].Trim()] = parts[1].Trim();
                }
            }
            catch { }
        }
        
        private string NormalizeAbbreviations(string institutionName)
        {
            if (_institutionMappings.TryGetValue(institutionName, out string mappedName)) return mappedName;
            return institutionName;
        }

        public IReadOnlyList<string> GetAllInstitutionNames() => _institutionNames.AsReadOnly();
    }
}
