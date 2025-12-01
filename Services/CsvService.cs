
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
        
        private static readonly HashSet<string> RightAlignedColumns = new HashSet<string>
        {
            "THERanking", "OverallGradeGPA", "DegreeStatus", "UKGrade"
        };
        
        private static readonly HashSet<string> DateColumns = new HashSet<string>
        {
            "ReceivedDate", "DueDate", "DateOfBirth"
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
                throw new InvalidOperationException($"Cannot read Document 1 (Exported new applicants inTray file).\n\nPlease close the file if it's open in Excel or another program.\n\nFile: {Path.GetFileName(filePath)}", ioEx);
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
                throw new InvalidOperationException($"Cannot read Document 2 (Dept. Application Reports file).\n\nPlease close the file if it's open in Excel or another program.\n\nFile: {Path.GetFileName(filePath)}", ioEx);
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
                    
                    // Write headers with neutral styling
                    for (int col = 0; col < ColumnOrder.Count; col++)
                    {
                        var headerCell = worksheet.Cells[1, col + 1];
                        headerCell.Value = ColumnOrder[col];
                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(240, 240, 240));
                        headerCell.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        headerCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                    
                    // Write data
                    int row = 2;
                    foreach (var record in records)
                    {
                        for (int col = 0; col < ColumnOrder.Count; col++)
                        {
                            string columnName = ColumnOrder[col];
                            string value = columnName switch
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
                            
                            var cell = worksheet.Cells[row, col + 1];
                            
                            // Apply alignment
                            if (RightAlignedColumns.Contains(columnName))
                            {
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            }
                            else
                            {
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            }
                            
                            // Apply specific formats
                            if (columnName == "StudentNo")
                            {
                                // Store as number to avoid "stored as text" warning
                                if (long.TryParse(value, out long studentNoValue))
                                {
                                    cell.Value = studentNoValue;
                                    cell.Style.Numberformat.Format = "0"; // Whole number, no decimals
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "THERanking")
                            {
                                // Store as text
                                cell.Style.Numberformat.Format = "@";
                                cell.Value = value;
                            }
                            else if (columnName == "Gender")
                            {
                                // Text format
                                cell.Style.Numberformat.Format = "@";
                                cell.Value = value;
                            }
                            else if (DateColumns.Contains(columnName) && !string.IsNullOrWhiteSpace(value))
                            {
                                // Date format
                                if (DateTime.TryParse(value, out DateTime dateValue))
                                {
                                    cell.Value = dateValue;
                                    cell.Style.Numberformat.Format = "dd/mm/yy";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "OverallGradeGPA" && !string.IsNullOrWhiteSpace(value))
                            {
                                // Percentage format
                                if (double.TryParse(value.Replace("%", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out double percentValue))
                                {
                                    cell.Value = percentValue / 100.0; // Excel percentages are 0-1
                                    cell.Style.Numberformat.Format = "0%";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "UKGrade" && !string.IsNullOrWhiteSpace(value))
                            {
                                // Number format with 1 decimal place
                                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double gradeValue))
                                {
                                    cell.Value = gradeValue;
                                    cell.Style.Numberformat.Format = "0.0";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else
                            {
                                cell.Value = value;
                            }
                        }
                        row++;
                    }
                    
                    // Suppress "number stored as text" warnings for THERanking column
                    int theRankingCol = ColumnOrder.IndexOf("THERanking") + 1;
                    if (row > 2) // Only if we have data rows
                    {
                        var theRankingRange = worksheet.Cells[2, theRankingCol, row - 1, theRankingCol];
                        var ignoredError = worksheet.IgnoredErrors.Add(theRankingRange);
                        ignoredError.NumberStoredAsText = true;
                    }
                    
                    // Auto-fit columns with custom widths
                    for (int col = 1; col <= ColumnOrder.Count; col++)
                    {
                        string columnName = ColumnOrder[col - 1];
                        
                        if (columnName == "EquivalencyNote")
                        {
                            worksheet.Column(col).Width = 18;
                        }
                        else if (columnName == "InstitutionName")
                        {
                            worksheet.Column(col).Width = 48;
                        }
                        else
                        {
                            worksheet.Column(col).AutoFit();
                        }
                    }
                    
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