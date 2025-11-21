// Services/EquivalencyService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ADMerger.Models;
using CsvHelper;
using CsvHelper.Configuration;

namespace ADMerger.Services
{
    public class EquivalencyService : IEquivalencyService
    {
        private readonly Dictionary<string, DegreeEquivalency> _equivalencies = new Dictionary<string, DegreeEquivalency>();
        
        public int Count => _equivalencies.Count;
        
        public void LoadEquivalencies()
        {
            try
            {
                StreamReader reader = null;
                string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "ucl_degree_equivalencies_FINAL.csv");
                
                if (File.Exists(csvPath))
                {
                    reader = new StreamReader(csvPath);
                }
                else
                {
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceName = "ADMerger.data.ucl_degree_equivalencies_FINAL.csv";
                    var stream = assembly.GetManifestResourceStream(resourceName);
                    
                    if (stream == null)
                        throw new FileNotFoundException("Equivalencies data not found.");
                    
                    reader = new StreamReader(stream);
                }
                
                using (reader)
                {
                    var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        HeaderValidated = null,
                        MissingFieldFound = null,
                        Delimiter = "\t"
                    };
                    
                    using var csv = new CsvReader(reader, config);
                    csv.Read();
                    csv.ReadHeader();
                    
                    while (csv.Read())
                    {
                        try
                        {
                            string country = csv.GetField(0)?.Trim().TrimStart('\'');
                            string third = csv.GetField(1)?.Trim().TrimStart('\'').TrimStart('<');
                            string secondLower = csv.GetField(2)?.Trim().TrimStart('\'');
                            string secondUpper = csv.GetField(3)?.Trim().TrimStart('\'');
                            string first = csv.GetField(4)?.Trim().TrimStart('\'');
                            
                            if (!string.IsNullOrWhiteSpace(country))
                            {
                                _equivalencies[country] = new DegreeEquivalency
                                {
                                    Country = country,
                                    Third = third,
                                    SecondLower = secondLower,
                                    SecondUpper = secondUpper,
                                    First = first
                                };
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Could not load equivalencies: {ex.Message}", ex);
            }
        }
        
        public DegreeEquivalency GetEquivalency(string country)
        {
            if (string.IsNullOrWhiteSpace(country))
                return null;
            
            return _equivalencies.ContainsKey(country.Trim()) ? _equivalencies[country.Trim()] : null;
        }
        
        public Dictionary<string, DegreeEquivalency> GetAllEquivalencies()
        {
            return new Dictionary<string, DegreeEquivalency>(_equivalencies);
        }
    }
}