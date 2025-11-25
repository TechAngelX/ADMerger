using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ADMerger.Models;

namespace ADMerger.Services
{
    public class OutputService
    {
        private readonly Dictionary<string, string> programmeMapping = new Dictionary<string, string>
        {
            {"MSc Artificial Intelligence for Biomedicine and Healthcare", "AIBH"},
            {"MSc Artificial Intelligence for Sustainable Development", "AISD"},
            {"MSc Artificial Intelligence and Data Engineering", "AIDE"},
            {"MSc Information Security", "ISEC"},
            {"MSc Computational Finance", "CF"},
            {"MSc Financial Risk Management", "FRM"},
            {"MSc Financial Technology", "FT"},
            {"MSc Emerging Digital Technologies", "EDT"},
            {"MSc Machine Learning", "ML"},
            {"MSc Data Science and Machine Learning", "DSML"},
            {"MSc Computational Statistics and Machine Learning", "CSML"},
            {"MSc Robotics and Artificial Intelligence", "RAI"},
            {"MSc Systems Engineering for the Internet of Things", "SEIOT"},
            {"MSc Disability, Design and Innovation", "DDI"},
            {"MSc Computer Science", "CS"},
            {"MSc Software Systems Engineering", "SSE"},
            {"MSc Computer Graphics, Vision and Imaging", "CGVI"}
        };

        public string GetProgrammeCode(string programmeName)
        {
            return programmeMapping.ContainsKey(programmeName) ? programmeMapping[programmeName] : programmeName;
        }

        public string GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath)
        {
            var programmeGroups = data.GroupBy(record => record.Programme).ToList();
            var outputPaths = new List<string>();
            
            var columnOrder = new List<string>
            {
                "ReceivedDate", "DueDate", "StudentNo", "Programme", "Forename", "Surname",
                "Gender", "DateOfBirth", "FeeStatus", "CountryOfNationality", "QualificationName",
                "DegreeSubject", "InstitutionName", "THERanking", "CountryOfStudy",
                "EquivalencyNote", "OverallGradeGPA", "UKGrade", "Decision", "AT",
                "Note", "Progr. Adm", "Comment"
            };
            
            foreach (var group in programmeGroups)
            {
                var programme = group.Key;
                var records = group.ToList();
                
                var outputPath = Path.Combine(
                    outputFolderPath, 
                    programme + "_Latest_" + DateTime.Now.ToString("dd_MMM_yyyy_HHmm") + ".csv");
                
                using var writer = new StreamWriter(outputPath);
                writer.WriteLine(string.Join(",", columnOrder));
                
                foreach (var record in records)
                {
                    var values = columnOrder.Select(column => GetFieldValue(record, column)).ToList();
                    writer.WriteLine(string.Join(",", values.Select(EscapeCSV)));
                }
                
                outputPaths.Add(outputPath);
            }
            
            return string.Join("\n", outputPaths);
        }

        private string GetFieldValue(OutputRecord record, string column)
        {
            return column switch
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
        }

        private string EscapeCSV(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        public string FormatDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return "";
            
            string[] formats = {
                "dd-MMM-yy", "dd-MMM-yyyy", "dd/MM/yy", "dd/MM/yyyy",
                "yyyy-MM-dd", "MM/dd/yyyy"
            };
            
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(dateStr, format, System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, out DateTime date))
                {
                    return date.ToString("dd/MM/yy");
                }
            }
            
            if (DateTime.TryParse(dateStr, out DateTime generalDate))
                return generalDate.ToString("dd/MM/yy");
            
            return dateStr;
        }

        public string CalculateDueDate(string receivedDateStr)
        {
            if (string.IsNullOrWhiteSpace(receivedDateStr)) return "";
            
            string[] formats = {
                "dd-MMM-yy", "dd-MMM-yyyy", "dd/MM/yy", "dd/MM/yyyy",
                "yyyy-MM-dd", "MM/dd/yyyy"
            };
            
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(receivedDateStr, format, System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, out DateTime date))
                {
                    return date.AddDays(42).ToString("dd/MM/yy");
                }
            }
            
            if (DateTime.TryParse(receivedDateStr, out DateTime generalDate))
                return generalDate.AddDays(42).ToString("dd/MM/yy");
            
            return "";
        }
    }
}
