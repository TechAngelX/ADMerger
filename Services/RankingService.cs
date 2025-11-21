// Services/RankingService.cs

using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ADMerger.Services
{
    public class RankingService : IRankingService
    {
        private readonly Dictionary<string, string> _rankings = new Dictionary<string, string>();
        private readonly List<string> _institutionNames = new List<string>();
        private readonly IInstitutionMatchingService _matchingService;
        
        public int Count => _rankings.Count;
        
        public RankingService(IInstitutionMatchingService matchingService)
        {
            _matchingService = matchingService ?? throw new ArgumentNullException(nameof(matchingService));
        }
        
        public void LoadRankings()
        {
            try
            {
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "THE Ranking 2026.xlsx");
                
                if (!File.Exists(excelPath))
                    throw new FileNotFoundException("THE Rankings file not found.");
                
                using var package = new ExcelPackage(new FileInfo(excelPath));
                var worksheet = package.Workbook.Worksheets[0];
                
                if (worksheet.Dimension == null)
                    throw new InvalidOperationException("THE Rankings sheet is empty.");
                
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
                            _rankings[institutionName] = rank;
                            _institutionNames.Add(institutionName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Could not load THE Rankings: {ex.Message}", ex);
            }
        }
        
        public string GetRanking(string institutionName)
        {
            if (string.IsNullOrWhiteSpace(institutionName))
                return "NR";
            
            if (_rankings.ContainsKey(institutionName))
                return _rankings[institutionName];
            
            string bestMatch = _matchingService.FindBestMatch(institutionName, _institutionNames);
            
            if (bestMatch != null)
                return _rankings[bestMatch];
            
            return "NR";
        }
        
        public IReadOnlyList<string> GetAllInstitutionNames()
        {
            return _institutionNames.AsReadOnly();
        }
    }
}