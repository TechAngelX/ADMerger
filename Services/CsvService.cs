// Services/CsvService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ADMerger.Models;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ADMerger.Services
{
    public class CsvService : ICsvService
    {
        private static readonly List<string> ColumnOrder = new List<string>
        {
            "ReceivedDate", "DueDate", "StudentNo", "Programme", "Forename", "Surname",
            "Gender", "DateOfBirth", "FeeStatus", "CountryOfNationality", "QualificationName",
            "DegreeSubject", "InstitutionName", "THERanking", "CountryOfStudy",
            "EquivalencyNote", "OverallGradeGPA", "DegreeStatus", "UKGrade", "Decision", "AT", "Note",
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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            var programmeGroups = data.GroupBy(record => record.Programme).ToList();
            var outputPaths = new List<string>();
            
            foreach (var group in programmeGroups)
            {
                var programme = group.Key;
                var records = group.OrderBy(r => ParseDate(r.ReceivedDate)).ToList();
                
                var outputPath = Path.Combine(
                    outputFolderPath, 
                    programme + "_Latest_" + DateTime.Now.ToString("dd_MMM_yyyy_HHmm") + ".xlsx");
                
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add(programme);
                    
                    // Write headers
                    for (int col = 0; col < ColumnOrder.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = ColumnOrder[col];
                        worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                        worksheet.Cells[1, col + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, col + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(59, 130, 246));
                        worksheet.Cells[1, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        worksheet.Cells[1, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    
                    // Write data
                    int row = 2;
                    foreach (var record in records)
                    {
                        for (int col = 0; col < ColumnOrder.Count; col++)
                        {
                            string value = ColumnOrder[col] switch
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
                                "DegreeStatus" => record.DegreeStatus ?? "",
                                "UKGrade" => record.UKGrade ?? "",
                                _ => ""
                            };
                            
                            worksheet.Cells[row, col + 1].Value = value;
                            worksheet.Cells[row, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        row++;
                    }
                    
                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    
                    package.SaveAs(new FileInfo(outputPath));
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
