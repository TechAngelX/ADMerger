using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using ADMerger.Models;

namespace ADMerger.Services
{
    public class DataService
    {
        public List<InTrayRecord> LoadInTrayData(string filePath)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HeaderValidated = null,
                MissingFieldFound = null
            };
            
            using var reader = new StringReader(File.ReadAllText(filePath));
            using var csv = new CsvReader(reader, config);
            return csv.GetRecords<InTrayRecord>().ToList();
        }

        public List<ApplicationRecord> LoadApplicationData(string filePath)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HeaderValidated = null,
                MissingFieldFound = null
            };
            
            using var reader = new StringReader(File.ReadAllText(filePath));
            using var csv = new CsvReader(reader, config);
            return csv.GetRecords<ApplicationRecord>().ToList();
        }

        public Dictionary<string, DegreeEquivalency> LoadDegreeEquivalencies()
        {
            var equivalencies = new Dictionary<string, DegreeEquivalency>();
            
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
                        return equivalencies;
                    
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
                                equivalencies[country] = new DegreeEquivalency
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
            catch { }
            
            return equivalencies;
        }
    }
}
