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
            "FeeStatus", "QualificationName",
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
            "ReceivedDate", "DueDate"
        };

        // Additional field property names that should be formatted as dates
        private static readonly HashSet<string> AdditionalDateFields = new HashSet<string>
        {
            "DateOfBirth",
            "QualificationEndDate",
            "AdmissionsReferredToDepartmentDate",
            "DepartmentRecommendedDecisionDate",
            "DepositDueDate",
            "InitialDecisionDate",
            "ReplyByDate"
        };

        public List<InTrayRecord> LoadInTrayRecords(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".xlsx")
                {
                    return LoadInTrayRecordsFromExcel(filePath);
                }
                
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

        private List<InTrayRecord> LoadInTrayRecordsFromExcel(string filePath)
        {
            var records = new List<InTrayRecord>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null) return records;

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text.Trim();
                    if (!string.IsNullOrEmpty(header))
                    {
                        headerMap[header] = col;
                    }
                }

                string GetValue(int row, string headerName)
                {
                    if (headerMap.TryGetValue(headerName, out int colIndex))
                    {
                        var cell = worksheet.Cells[row, colIndex];
                        return cell.Text?.Trim() ?? "";
                    }
                    return "";
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new InTrayRecord
                    {
                        ReceivedOn = GetValue(row, "Received on"),
                        StudentNo = GetValue(row, "Student No."),
                        Name = GetValue(row, "Name")
                    };
                    records.Add(record);
                }
            }

            return records;
        }
        
        public List<ApplicationRecord> LoadApplicationRecords(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".xlsx")
                {
                    return LoadApplicationRecordsFromExcel(filePath);
                }
                
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

        private List<ApplicationRecord> LoadApplicationRecordsFromExcel(string filePath)
        {
            var records = new List<ApplicationRecord>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null) return records;

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text.Trim();
                    if (!string.IsNullOrEmpty(header))
                    {
                        headerMap[header] = col;
                    }
                }

                string GetValue(int row, string headerName)
                {
                    if (headerMap.TryGetValue(headerName, out int colIndex))
                    {
                        return worksheet.Cells[row, colIndex].Text?.Trim() ?? "";
                    }
                    return "";
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new ApplicationRecord
                    {
                        ApplicantID = GetValue(row, "Applicant ID"),
                        Programme = GetValue(row, "Programme"),
                        Forename = GetValue(row, "Forename"),
                        Surname = GetValue(row, "Surname"),
                        FeeStatus = GetValue(row, "Fee Status"),
                        QualificationName = GetValue(row, "Qualification name"),
                        DegreeSubject = GetValue(row, "Degree subject"),
                        InstitutionName = GetValue(row, "Institution name"),
                        CountryOfStudy = GetValue(row, "Country of study"),
                        OverallGradeGPA = GetValue(row, "Overall  grade/GPA"),
                        EquivalencyNote = GetValue(row, "Equivalency note"),
                        GradeAchievedPending = GetValue(row, "Grade Achieved/Pending"),

                        // Additional optional fields
                        KnownAs = GetValue(row, "Known as"),
                        ModeOfAttendance = GetValue(row, "Mode of attendance"),
                        Location = GetValue(row, "Location"),
                        State = GetValue(row, "State"),
                        EmailAddress = GetValue(row, "Email address"),
                        Gender = GetValue(row, "Gender"),
                        DateOfBirth = GetValue(row, "Date of Birth"),
                        CountryOfNationality = GetValue(row, "Country of Nationality"),
                        QualificationEndDate = GetValue(row, "Qualification end date"),
                        TotalMarkEquivalency = GetValue(row, "Total Mark equivalency"),
                        AdmissionsReferralNote = GetValue(row, "Admissions referral note"),
                        AdmissionsReferredToDepartmentDate = GetValue(row, "Admissions referred to department date"),
                        DepartmentRecommendedDecision = GetValue(row, "Department recommended decision"),
                        DepartmentRecommendedDecisionDate = GetValue(row, "Department recommended decision date"),
                        DepositDueDate = GetValue(row, "Deposit due date"),
                        DepositPaymentStatus = GetValue(row, "Deposit payment status"),
                        InitialDecision = GetValue(row, "Initial decision"),
                        InitialDecisionDate = GetValue(row, "Initial Decision date"),
                        DecisionResponse = GetValue(row, "Decision/Response"),
                        ReplyByDate = GetValue(row, "Reply by date"),
                        AcademicYear = GetValue(row, "Academic year"),
                        Tag = GetValue(row, "Tag"),
                        ELPType = GetValue(row, "ELP type"),
                        ELPVerificationStatus = GetValue(row, "ELP verification status")
                    };
                    records.Add(record);
                }
            }

            return records;
        }
        
        public List<string> GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath, List<string>? additionalFields = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Build dynamic column order: standard columns + manual entry columns + additional fields at the end
            var dynamicColumnOrder = new List<string>(ColumnOrder);

            // Add additional fields at the very end (after "Comment")
            if (additionalFields != null && additionalFields.Count > 0)
            {
                dynamicColumnOrder.AddRange(additionalFields);
            }

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

                    for (int col = 0; col < dynamicColumnOrder.Count; col++)
                    {
                        var headerCell = worksheet.Cells[1, col + 1];
                        var columnName = dynamicColumnOrder[col];
                        headerCell.Value = columnName;
                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;

                        // Check if this is an additional field (not in original ColumnOrder)
                        bool isAdditionalField = additionalFields != null && additionalFields.Contains(columnName);

                        if (isAdditionalField)
                        {
                            // Highlight additional fields with a different color (light blue)
                            headerCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 230, 241));
                            headerCell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(31, 78, 121));
                        }
                        else
                        {
                            // Standard header color (light gray)
                            headerCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(240, 240, 240));
                            headerCell.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        }

                        headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        headerCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                    
                    int row = 2;
                    foreach (var record in records)
                    {
                        for (int col = 0; col < dynamicColumnOrder.Count; col++)
                        {
                            string columnName = dynamicColumnOrder[col];
                            string value = columnName switch
                            {
                                "ReceivedDate" => record.ReceivedDate ?? "",
                                "DueDate" => record.DueDate ?? "",
                                "StudentNo" => record.StudentNo ?? "",
                                "Programme" => record.Programme ?? "",
                                "Forename" => record.Forename ?? "",
                                "Surname" => record.Surname ?? "",
                                "FeeStatus" => record.FeeStatus ?? "",
                                "QualificationName" => record.QualificationName ?? "",
                                "DegreeSubject" => record.DegreeSubject ?? "",
                                "InstitutionName" => record.InstitutionName ?? "",
                                "THERanking" => record.THERanking ?? "NR",
                                "CountryOfStudy" => record.CountryOfStudy ?? "",
                                "EquivalencyNote" => record.EquivalencyNote ?? "",
                                "OverallGradeGPA" => record.OverallGradeGPA ?? "",
                                "DegreeStatus" => record.DegreeStatus ?? "",
                                "UKGrade" => record.UKGrade ?? "",
                                _ => record.AdditionalFieldValues.ContainsKey(columnName) ? record.AdditionalFieldValues[columnName] ?? "" : ""
                            };
                            
                            var cell = worksheet.Cells[row, col + 1];
                            
                            if (RightAlignedColumns.Contains(columnName))
                            {
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            }
                            else
                            {
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            }
                            
                            if (columnName == "StudentNo")
                            {
                                if (long.TryParse(value, out long studentNoValue))
                                {
                                    cell.Value = studentNoValue;
                                    cell.Style.Numberformat.Format = "0";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "THERanking")
                            {
                                cell.Style.Numberformat.Format = "@";
                                cell.Value = value;
                            }
                            else if (DateColumns.Contains(columnName) && !string.IsNullOrWhiteSpace(value))
                            {
                                if (DateTime.TryParse(value, out DateTime dateValue))
                                {
                                    cell.Value = dateValue;
                                    cell.Style.Numberformat.Format = "dd/mm/yyyy";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (AdditionalDateFields.Contains(columnName) && !string.IsNullOrWhiteSpace(value))
                            {
                                // Format additional date fields (DateOfBirth, QualificationEndDate, etc.)
                                if (DateTime.TryParse(value, out DateTime dateValue))
                                {
                                    cell.Value = dateValue;
                                    cell.Style.Numberformat.Format = "dd/mm/yyyy";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "OverallGradeGPA" && !string.IsNullOrWhiteSpace(value))
                            {
                                if (double.TryParse(value.Replace("%", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out double percentValue))
                                {
                                    cell.Value = percentValue / 100.0;
                                    cell.Style.Numberformat.Format = "0%";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "UKGrade" && !string.IsNullOrWhiteSpace(value))
                            {
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
                    
                    int theRankingCol = ColumnOrder.IndexOf("THERanking") + 1;
                    if (row > 2)
                    {
                        var theRankingRange = worksheet.Cells[2, theRankingCol, row - 1, theRankingCol];
                        var ignoredError = worksheet.IgnoredErrors.Add(theRankingRange);
                        ignoredError.NumberStoredAsText = true;
                    }
                    
                    for (int col = 1; col <= dynamicColumnOrder.Count; col++)
                    {
                        string columnName = dynamicColumnOrder[col - 1];
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
            
            return outputPaths;
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
