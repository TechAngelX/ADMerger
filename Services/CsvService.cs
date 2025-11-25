// Services/CsvService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ADMerger.Models;
using CsvHelper;
using CsvHelper.Configuration;

namespace ADMerger.Services
{
    public class CsvService : ICsvService
    {
        private static readonly List<string> ColumnOrder = new List<string>
        {
            "ReceivedDate", "DueDate", "StudentNo", "Programme", "Forename", "Surname",
            "Gender", "DateOfBirth", "FeeStatus", "CountryOfNationality", "QualificationName",
            "DegreeSubject", "InstitutionName", "THERanking", "CountryOfStudy",
            "EquivalencyNote", "OverallGradeGPA", "UKGrade", "Decision", "AT", "Note",
            "Progr. Adm", "Comment"
        };
        
        public List<InTrayRecord> LoadInTrayRecords(string filePath)
        {
            try
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
            catch (IOException ioEx)
            {
                throw new InvalidOperationException($"Cannot read Document 1 (Exported new applicants csv file).\n\nPlease close the file if it's open in Excel or another program.\n\nFile: {Path.GetFileName(filePath)}", ioEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error loading InTray records: {ex.Message}", ex);
            }
        }
        
        public List<ApplicationRecord> LoadApplicationRecords(string filePath)
        {
            try
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
            catch (IOException ioEx)
            {
                throw new InvalidOperationException($"Cannot read Document 2 (the PorticoDepartment Application Reports file).\n\nPlease close the file if it's open in Excel or another program.\n\nFile: {Path.GetFileName(filePath)}", ioEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error loading Application records: {ex.Message}", ex);
            }
        }
        
        public string GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath)
        {
            var programmeGroups = data.GroupBy(record => record.Programme).ToList();
            var outputPaths = new List<string>();
            
            foreach (var group in programmeGroups)
            {
                var programme = group.Key;
                var records = group.OrderBy(r => ParseDate(r.ReceivedDate)).ToList();
                
                var outputPath = Path.Combine(
                    outputFolderPath, 
                    programme + "_Latest_" + DateTime.Now.ToString("dd_MMM_yyyy_HHmm") + ".csv");
                
                using var writer = new StreamWriter(outputPath);
                writer.WriteLine(string.Join(",", ColumnOrder));
                
                foreach (var record in records)
                {
                    var values = new List<string>();
                    
                    foreach (var column in ColumnOrder)
                    {
                        string value = column switch
                        {
                            "ReceivedDate" => record.ReceivedDate ?? "",
                            "DueDate" => record.DueDate ?? "",
                            "StudentNo" => record.StudentNo ?? "",
                            "Programme" => record.Programme ?? "",
                            "Forename" => record.Forename ?? "",
                            "Surname" => record.Surname ?? "",
                            "Gender" => record.Gender ?? "",
                            "DateOfBirth" => record.DateOfBirth ?? "",
                            "FeeStatus" => record.FeeStatus ?? "",
                            "CountryOfNationality" => record.CountryOfNationality ?? "",
                            "QualificationName" => record.QualificationName ?? "",
                            "DegreeSubject" => record.DegreeSubject ?? "",
                            "InstitutionName" => record.InstitutionName ?? "",
                            "THERanking" => record.THERanking ?? "NR",
                            "CountryOfStudy" => record.CountryOfStudy ?? "",
                            "EquivalencyNote" => record.EquivalencyNote ?? "",
                            "OverallGradeGPA" => record.OverallGradeGPA ?? "",
                            "UKGrade" => record.UKGrade ?? "",
                            _ => ""
                        };
                        
                        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                        {
                            value = "\"" + value.Replace("\"", "\"\"") + "\"";
                        }
                        
                        values.Add(value);
                    }
                    
                    writer.WriteLine(string.Join(",", values));
                }
                
                outputPaths.Add(outputPath);
            }
            
            return string.Join("\n", outputPaths);
        }
        
        private DateTime ParseDate(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return DateTime.MinValue;
            
            if (DateTime.TryParseExact(dateString, "dd/MM/yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                return result;
            
            if (DateTime.TryParse(dateString, out result))
                return result;
            
            return DateTime.MinValue;
        }
    }
}

