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
            LoadInstitutionMappings();
            
            try
            {
                Stream excelStream = null;
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "THE Ranking 2026.xlsx");
                
                if (File.Exists(excelPath))
                {
                    excelStream = File.OpenRead(excelPath);
                }
                else
                {
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceName = "ADMerger.data.THE_Ranking_2026.xlsx";
                    excelStream = assembly.GetManifestResourceStream(resourceName);
                    
                    if (excelStream == null)
                        throw new FileNotFoundException("THE Rankings file not found as file or embedded resource.");
                }
                
                using (excelStream)
                using (var package = new ExcelPackage(excelStream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    
                    if (worksheet.Dimension == null)
                        throw new InvalidOperationException("THE Rankings sheet is empty.");
                    
                    int totalRows = worksheet.Dimension.Rows;
                    int loadedCount = 0;
                    
                    for (int row = 2; row <= totalRows; row++)
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
                                loadedCount++;
                            }
                        }
                    }
                    
                    if (loadedCount == 0)
                        throw new InvalidOperationException($"No data loaded from Excel. Total rows: {totalRows}");
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
            
            string originalName = institutionName;
            institutionName = NormalizeAbbreviations(institutionName);
            
            if (institutionName == "NOT_RANKED")
            {
                LogToFile($"NOT RANKED: '{originalName}' (marked as not in THE Rankings)");
                return "NR";
            }
            
            LogToFile($"Original: '{originalName}' | Normalized: '{institutionName}' | InDict: {_rankings.ContainsKey(institutionName)}");
            
            if (_rankings.ContainsKey(institutionName))
            {
                string rank = CleanRankingText(_rankings[institutionName]);
                LogToFile($"EXACT MATCH: '{institutionName}' = Rank {rank}");
                return rank;
            }
            
            string bestMatch = _matchingService.FindBestMatch(institutionName, _institutionNames);
            
            if (bestMatch != null)
            {
                string rank = CleanRankingText(_rankings[bestMatch]);
                LogToFile($"FUZZY MATCH: '{institutionName}' -> '{bestMatch}' = Rank {rank}");
                return rank;
            }
            
            LogToFile($"NO MATCH: '{institutionName}'");
            return "NR";
        }
        
        private void LogToFile(string message)
        {
            try
            {
                string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ranking_matches.log");
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { }
        }
        
        private string CleanRankingText(string ranking)
        {
            if (string.IsNullOrWhiteSpace(ranking))
                return "NR";
            
            byte[] bytes = Encoding.UTF8.GetBytes(ranking);
            string cleanRanking = Encoding.UTF8.GetString(bytes);
            
            cleanRanking = cleanRanking.Replace(char.ConvertFromUtf32(0x2013), "-");
            cleanRanking = cleanRanking.Replace(char.ConvertFromUtf32(0x2014), "-");
            
            if (cleanRanking.Contains("â") || cleanRanking.Contains("€"))
            {
                cleanRanking = System.Text.RegularExpressions.Regex.Replace(cleanRanking, @"â€.", "-");
            }
            
            return cleanRanking;
        }
        
        private void LoadInstitutionMappings()
        {
            try
            {
                string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "institution_mappings.csv");
                
                if (!File.Exists(csvPath))
                {
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceName = "ADMerger.data.institution_mappings.csv";
                    var stream = assembly.GetManifestResourceStream(resourceName);
                    
                    if (stream == null)
                        return;
                    
                    using (var reader = new StreamReader(stream))
                    {
                        reader.ReadLine();
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            var parts = line.Split(',');
                            if (parts.Length >= 2)
                            {
                                _institutionMappings[parts[0].Trim()] = parts[1].Trim();
                            }
                        }
                    }
                }
                else
                {
                    using (var reader = new StreamReader(csvPath))
                    {
                        reader.ReadLine();
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            var parts = line.Split(',');
                            if (parts.Length >= 2)
                            {
                                _institutionMappings[parts[0].Trim()] = parts[1].Trim();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }
        
        private string NormalizeAbbreviations(string institutionName)
        {
            if (_institutionMappings.TryGetValue(institutionName, out string mappedName))
            {
                return mappedName;
            }
            
            return institutionName;
        }
        
        public IReadOnlyList<string> GetAllInstitutionNames()
        {
            return _institutionNames.AsReadOnly();
        }
    }
}

